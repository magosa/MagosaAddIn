using System;
using System.Runtime.InteropServices;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// COM例外処理とログ出力の統一を提供するクラス
    /// </summary>
    public static class ComExceptionHandler
    {
        #region COM例外処理

        /// <summary>
        /// COM例外を統一的に処理する
        /// </summary>
        /// <param name="action">実行するアクション</param>
        /// <param name="operationName">操作名</param>
        /// <param name="throwOnError">エラー時に例外をスローするか</param>
        /// <returns>処理が成功した場合true</returns>
        public static bool HandleComOperation(Action action, string operationName, bool throwOnError = true)
        {
            try
            {
                action?.Invoke();
                LogDebug($"{operationName}: 成功");
                return true;
            }
            catch (COMException comEx)
            {
                string errorMessage = $"{operationName}中にCOM例外が発生しました";
                LogDebug($"{errorMessage}: HRESULT=0x{comEx.HResult:X8}, Message={comEx.Message}");

                if (throwOnError)
                {
                    throw new Exception($"{errorMessage}: {GetComErrorDescription(comEx)}");
                }
                return false;
            }
            catch (Exception ex)
            {
                string errorMessage = $"{operationName}中にエラーが発生しました";
                LogDebug($"{errorMessage}: {ex.Message}");

                if (throwOnError)
                {
                    throw new Exception($"{errorMessage}: {ex.Message}");
                }
                return false;
            }
        }

        /// <summary>
        /// COM例外を統一的に処理する（戻り値あり）
        /// </summary>
        /// <typeparam name="T">戻り値の型</typeparam>
        /// <param name="func">実行する関数</param>
        /// <param name="operationName">操作名</param>
        /// <param name="defaultValue">エラー時のデフォルト値</param>
        /// <param name="throwOnError">エラー時に例外をスローするか</param>
        /// <returns>処理結果またはデフォルト値</returns>
        public static T HandleComOperation<T>(Func<T> func, string operationName, T defaultValue = default(T), bool throwOnError = true)
        {
            try
            {
                var result = func != null ? func.Invoke() : defaultValue;
                LogDebug($"{operationName}: 成功");
                return result;
            }
            catch (COMException comEx)
            {
                string errorMessage = $"{operationName}中にCOM例外が発生しました";
                LogDebug($"{errorMessage}: HRESULT=0x{comEx.HResult:X8}, Message={comEx.Message}");

                if (throwOnError)
                {
                    throw new Exception($"{errorMessage}: {GetComErrorDescription(comEx)}");
                }
                return defaultValue;
            }
            catch (Exception ex)
            {
                string errorMessage = $"{operationName}中にエラーが発生しました";
                LogDebug($"{errorMessage}: {ex.Message}");

                if (throwOnError)
                {
                    throw new Exception($"{errorMessage}: {ex.Message}");
                }
                return defaultValue;
            }
        }

        /// <summary>
        /// COM例外の詳細説明を取得
        /// </summary>
        /// <param name="comEx">COM例外</param>
        /// <returns>エラーの詳細説明</returns>
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
                default:
                    return $"COM例外 (HRESULT: 0x{comEx.HResult:X8})";
            }
        }

        #endregion

        #region ログ出力

        /// <summary>
        /// デバッグログを統一フォーマットで出力
        /// </summary>
        /// <param name="message">ログメッセージ</param>
        /// <param name="level">ログレベル</param>
        public static void LogDebug(string message, LogLevel level = LogLevel.Info)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string levelStr = GetLogLevelString(level);
            string formattedMessage = $"[{timestamp}] [{levelStr}] MagosaAddIn: {message}";

            System.Diagnostics.Debug.WriteLine(formattedMessage);

            // 必要に応じてファイル出力やイベントログ出力を追加可能
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
        /// ログレベルの文字列表現を取得
        /// </summary>
        /// <param name="level">ログレベル</param>
        /// <returns>ログレベル文字列</returns>
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

        #region ユーザーメッセージ生成

        /// <summary>
        /// ユーザー向けエラーメッセージを生成
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <param name="ex">例外オブジェクト</param>
        /// <returns>ユーザー向けメッセージ</returns>
        public static string CreateUserErrorMessage(string operationName, Exception ex)
        {
            if (ex is COMException comEx)
            {
                return $"{operationName}中にエラーが発生しました。\n\n詳細: {GetComErrorDescription(comEx)}\n\nPowerPointを再起動してお試しください。";
            }
            else
            {
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
    /// ログレベル列挙型
    /// </summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }
}