using System;
using System.Runtime.InteropServices;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// COM例外処理とログ出力の統一を提供するクラス
    /// </summary>
    public static class ComExceptionHandler
    {
        #region COM例外処理（統一版）

        /// <summary>
        /// COM操作を安全に実行する（戻り値なし）
        /// </summary>
        /// <param name="action">実行するアクション</param>
        /// <param name="operationName">操作名</param>
        /// <param name="suppressErrors">エラーを抑制するか（false=例外をスロー、true=ログのみ出力）</param>
        /// <returns>処理が成功した場合true</returns>
        public static bool ExecuteComOperation(Action action, string operationName, bool suppressErrors = false)
        {
            if (action == null)
            {
                LogError($"{operationName}: アクションがnullです");
                return false;
            }

            try
            {
                action.Invoke();
                LogDebug($"{operationName}: 成功", LogLevel.Debug);
                return true;
            }
            catch (COMException comEx)
            {
                return HandleComExceptionInternal(operationName, comEx, suppressErrors);
            }
            catch (InvalidOperationException invEx)
            {
                return HandleInvalidOperationInternal(operationName, invEx, suppressErrors);
            }
            catch (ArgumentException argEx)
            {
                return HandleArgumentExceptionInternal(operationName, argEx, suppressErrors);
            }
            catch (Exception ex)
            {
                return HandleGeneralExceptionInternal(operationName, ex, suppressErrors);
            }
        }

        /// <summary>
        /// COM操作を安全に実行する（戻り値あり）
        /// </summary>
        /// <typeparam name="T">戻り値の型</typeparam>
        /// <param name="func">実行する関数</param>
        /// <param name="operationName">操作名</param>
        /// <param name="defaultValue">エラー時のデフォルト値</param>
        /// <param name="suppressErrors">エラーを抑制するか（false=例外をスロー、true=デフォルト値を返す）</param>
        /// <returns>処理結果またはデフォルト値</returns>
        public static T ExecuteComOperation<T>(Func<T> func, string operationName, T defaultValue = default(T), bool suppressErrors = false)
        {
            if (func == null)
            {
                LogError($"{operationName}: 関数がnullです");
                return defaultValue;
            }

            try
            {
                var result = func.Invoke();
                LogDebug($"{operationName}: 成功", LogLevel.Debug);
                return result;
            }
            catch (COMException comEx)
            {
                HandleComExceptionInternal(operationName, comEx, suppressErrors);
                return defaultValue;
            }
            catch (InvalidOperationException invEx)
            {
                HandleInvalidOperationInternal(operationName, invEx, suppressErrors);
                return defaultValue;
            }
            catch (ArgumentException argEx)
            {
                HandleArgumentExceptionInternal(operationName, argEx, suppressErrors);
                return defaultValue;
            }
            catch (Exception ex)
            {
                HandleGeneralExceptionInternal(operationName, ex, suppressErrors);
                return defaultValue;
            }
        }

        /// <summary>
        /// 旧メソッド（互換性維持）- 新規コードでは使用非推奨
        /// </summary>
        [Obsolete("ExecuteComOperation()を使用してください")]
        public static bool HandleComOperation(Action action, string operationName, bool throwOnError = true)
        {
            return ExecuteComOperation(action, operationName, suppressErrors: !throwOnError);
        }

        /// <summary>
        /// 旧メソッド（互換性維持）- 新規コードでは使用非推奨
        /// </summary>
        [Obsolete("ExecuteComOperation<T>()を使用してください")]
        public static T HandleComOperation<T>(Func<T> func, string operationName, T defaultValue = default(T), bool throwOnError = true)
        {
            return ExecuteComOperation(func, operationName, defaultValue, suppressErrors: !throwOnError);
        }

        #endregion

        #region 内部例外処理メソッド

        /// <summary>
        /// COM例外の内部処理
        /// </summary>
        private static bool HandleComExceptionInternal(string operationName, COMException comEx, bool suppressErrors)
        {
            string errorMessage = $"{operationName}中にCOM例外が発生しました";
            string detailMessage = GetComErrorDescription(comEx);

            LogError($"{errorMessage}: HRESULT=0x{comEx.HResult:X8}, {detailMessage}");

            if (!suppressErrors)
            {
                throw new ComOperationException($"{errorMessage}: {detailMessage}", comEx);
            }
            return false;
        }

        /// <summary>
        /// InvalidOperation例外の内部処理
        /// </summary>
        private static bool HandleInvalidOperationInternal(string operationName, InvalidOperationException invEx, bool suppressErrors)
        {
            string errorMessage = $"{operationName}中に無効な操作が実行されました";
            LogError($"{errorMessage}: {invEx.Message}");

            if (!suppressErrors)
            {
                throw new InvalidOperationException($"{errorMessage}: {invEx.Message}", invEx);
            }
            return false;
        }

        /// <summary>
        /// Argument例外の内部処理
        /// </summary>
        private static bool HandleArgumentExceptionInternal(string operationName, ArgumentException argEx, bool suppressErrors)
        {
            string errorMessage = $"{operationName}中に引数エラーが発生しました";
            LogError($"{errorMessage}: {argEx.Message}");

            if (!suppressErrors)
            {
                throw new ArgumentException($"{errorMessage}: {argEx.Message}", argEx);
            }
            return false;
        }

        /// <summary>
        /// 一般例外の内部処理
        /// </summary>
        private static bool HandleGeneralExceptionInternal(string operationName, Exception ex, bool suppressErrors)
        {
            string errorMessage = $"{operationName}中に予期しないエラーが発生しました";
            LogError($"{errorMessage}: {ex.GetType().Name} - {ex.Message}");

            if (!suppressErrors)
            {
                throw new Exception($"{errorMessage}: {ex.Message}", ex);
            }
            return false;
        }

        /// <summary>
        /// COM例外の詳細説明を取得
        /// </summary>
        private static string GetComErrorDescription(COMException comEx)
        {
            switch ((uint)comEx.HResult)
            {
                case 0x800A01A8: // VBA_E_OBJECTDELETED
                    return "オブジェクトが削除されています";
                case 0x80004005: // E_FAIL
                    return "操作が失敗しました";
                case 0x80070005: // E_ACCESSDENIED
                    return "アクセスが拒否されました";
                case 0x800401E3: // MK_E_UNAVAILABLE
                    return "オブジェクトが利用できません";
                case 0x80020009: // DISP_E_EXCEPTION
                    return "ディスパッチ例外が発生しました";
                case 0x8002000E: // DISP_E_PARAMNOTFOUND
                    return "パラメータが見つかりません";
                default:
                    return $"COM例外 (HRESULT: 0x{comEx.HResult:X8})";
            }
        }

        #endregion

        #region ログ出力（改良版）

        /// <summary>
        /// ログレベル設定（本番環境では Warning 以上のみ出力）
        /// </summary>
        public static LogLevel MinimumLogLevel { get; set; } = LogLevel.Debug;

        /// <summary>
        /// 統一されたログ出力
        /// </summary>
        /// <param name="message">ログメッセージ</param>
        /// <param name="level">ログレベル</param>
        public static void LogDebug(string message, LogLevel level = LogLevel.Info)
        {
            // ログレベルフィルタリング
            if (level < MinimumLogLevel)
                return;

            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string levelStr = GetLogLevelString(level);
            string formattedMessage = $"[{timestamp}] [{levelStr}] MagosaAddIn: {message}";

            System.Diagnostics.Debug.WriteLine(formattedMessage);

            // 本番環境では追加でファイル出力やイベントログ出力を実装可能
            if (level >= LogLevel.Error)
            {
                // 重要なエラーは追加ログ出力（将来実装）
                // WriteToErrorLog(formattedMessage);
            }
        }

        /// <summary>
        /// エラーログを出力
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="ex">例外オブジェクト</param>
        public static void LogError(string message, Exception ex = null)
        {
            string errorMessage = ex != null ? $"{message}: {ex.Message}" : message;
            LogDebug(errorMessage, LogLevel.Error);
        }

        /// <summary>
        /// 警告ログを出力
        /// </summary>
        /// <param name="message">警告メッセージ</param>
        public static void LogWarning(string message)
        {
            LogDebug(message, LogLevel.Warning);
        }

        /// <summary>
        /// 情報ログを出力
        /// </summary>
        /// <param name="message">情報メッセージ</param>
        public static void LogInfo(string message)
        {
            LogDebug(message, LogLevel.Info);
        }

        /// <summary>
        /// ログレベルの文字列表現を取得
        /// </summary>
        private static string GetLogLevelString(LogLevel level)
        {
            switch (level)
            {
                case LogLevel.Error: return "ERROR";
                case LogLevel.Warning: return "WARN ";
                case LogLevel.Info: return "INFO ";
                case LogLevel.Debug: return "DEBUG";
                default: return "INFO ";
            }
        }

        #endregion

        #region ユーザーメッセージ生成（改良版）

        /// <summary>
        /// ユーザー向けエラーメッセージを生成
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <param name="ex">例外オブジェクト</param>
        /// <returns>ユーザー向けメッセージ</returns>
        public static string CreateUserErrorMessage(string operationName, Exception ex)
        {
            switch (ex)
            {
                case ComOperationException comOpEx:
                    return $"{operationName}中にエラーが発生しました。\n\n詳細: {comOpEx.UserMessage}\n\nPowerPointを再起動してお試しください。";

                case COMException comEx:
                    return $"{operationName}中にエラーが発生しました。\n\n詳細: {GetComErrorDescription(comEx)}\n\nPowerPointを再起動してお試しください。";

                case ArgumentException argEx:
                    return $"{operationName}の設定値に問題があります。\n\n詳細: {argEx.Message}";

                case InvalidOperationException invEx:
                    return $"{operationName}を実行できませんでした。\n\n詳細: {invEx.Message}\n\n図形の選択状態を確認してください。";

                default:
                    return $"{operationName}中にエラーが発生しました。\n\n詳細: {ex.Message}";
            }
        }

        /// <summary>
        /// 操作成功メッセージを生成
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <param name="details">詳細情報</param>
        /// <returns>成功メッセージ</returns>
        public static string CreateSuccessMessage(string operationName, string details = null)
        {
            return string.IsNullOrEmpty(details) ?
                $"{operationName}が完了しました。" :
                $"{operationName}が完了しました。{details}";
        }

        #endregion
    }

    /// <summary>
    /// ログレベル列挙型（改良版）
    /// </summary>
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warning = 2,
        Error = 3
    }

    /// <summary>
    /// COM操作専用例外クラス
    /// </summary>
    public class ComOperationException : Exception
    {
        public string UserMessage { get; }

        public ComOperationException(string message, Exception innerException = null)
            : base(message, innerException)
        {
            UserMessage = message;
        }

        public ComOperationException(string message, string userMessage, Exception innerException = null)
            : base(message, innerException)
        {
            UserMessage = userMessage;
        }
    }
}