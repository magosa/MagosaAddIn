using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// エラーハンドリングと入力検証の統一を提供するクラス
    /// </summary>
    public static class ErrorHandler
    {
        #region 入力検証

        /// <summary>
        /// 図形リストの検証
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="minimumCount">必要最小数</param>
        /// <param name="operationName">操作名</param>
        public static void ValidateShapes(List<PowerPoint.Shape> shapes, int minimumCount, string operationName)
        {
            if (shapes == null)
            {
                var message = CreateValidationMessage(operationName, "図形が選択されていません。");
                ComExceptionHandler.LogError($"{operationName}入力検証エラー", new ArgumentNullException(nameof(shapes)));
                throw new ArgumentNullException(nameof(shapes), message);
            }

            if (shapes.Count < minimumCount)
            {
                var message = CreateValidationMessage(operationName,
                    $"{minimumCount}つ以上の図形が必要です。現在の選択数: {shapes.Count}個");
                ComExceptionHandler.LogWarning($"{operationName}: 図形数不足 ({shapes.Count}/{minimumCount})");
                throw new ArgumentException(message);
            }

            ComExceptionHandler.LogDebug($"{operationName}: 図形検証成功 ({shapes.Count}個)");
        }

        /// <summary>
        /// 数値範囲の検証
        /// </summary>
        /// <param name="value">検証する値</param>
        /// <param name="min">最小値</param>
        /// <param name="max">最大値</param>
        /// <param name="parameterName">パラメータ名</param>
        /// <param name="operationName">操作名</param>
        public static void ValidateRange(int value, int min, int max, string parameterName, string operationName)
        {
            if (value < min || value > max)
            {
                var message = CreateValidationMessage(operationName,
                    $"{parameterName}は{min}以上{max}以下である必要があります。入力値: {value}");
                ComExceptionHandler.LogError($"{operationName}範囲検証エラー",
                    new ArgumentOutOfRangeException(parameterName, value, message));
                throw new ArgumentOutOfRangeException(parameterName, value, message);
            }

            ComExceptionHandler.LogDebug($"{operationName}: {parameterName}検証成功 ({value})");
        }

        /// <summary>
        /// 数値範囲の検証（float版）
        /// </summary>
        /// <param name="value">検証する値</param>
        /// <param name="min">最小値</param>
        /// <param name="max">最大値</param>
        /// <param name="parameterName">パラメータ名</param>
        /// <param name="operationName">操作名</param>
        public static void ValidateRange(float value, float min, float max, string parameterName, string operationName)
        {
            if (value < min || value > max)
            {
                var message = CreateValidationMessage(operationName,
                    $"{parameterName}は{min:F1}以上{max:F1}以下である必要があります。入力値: {value:F1}");
                ComExceptionHandler.LogError($"{operationName}範囲検証エラー",
                    new ArgumentOutOfRangeException(parameterName, value, message));
                throw new ArgumentOutOfRangeException(parameterName, value, message);
            }

            ComExceptionHandler.LogDebug($"{operationName}: {parameterName}検証成功 ({value:F1})");
        }

        /// <summary>
        /// セルサイズの検証
        /// </summary>
        /// <param name="cellWidth">セル幅</param>
        /// <param name="cellHeight">セル高さ</param>
        /// <param name="operationName">操作名</param>
        public static void ValidateCellSize(float cellWidth, float cellHeight, string operationName)
        {
            if (cellWidth <= Constants.MIN_CELL_SIZE || cellHeight <= Constants.MIN_CELL_SIZE)
            {
                var message = CreateValidationMessage(operationName,
                    $"セルサイズが小さすぎます。\n" +
                    $"計算されたサイズ: {cellWidth:F1}×{cellHeight:F1}pt\n" +
                    $"最小サイズ: {Constants.MIN_CELL_SIZE}pt\n" +
                    $"マージンを小さくするか、分割数を減らしてください。");
                ComExceptionHandler.LogError($"{operationName}セルサイズ検証エラー",
                    new ArgumentException(message));
                throw new ArgumentException(message);
            }

            ComExceptionHandler.LogDebug($"{operationName}: セルサイズ検証成功 ({cellWidth:F1}×{cellHeight:F1})");
        }

        /// <summary>
        /// 座標の検証
        /// </summary>
        /// <param name="x">X座標</param>
        /// <param name="y">Y座標</param>
        /// <param name="operationName">操作名</param>
        /// <returns>座標が有効な場合true</returns>
        public static bool ValidateCoordinates(float x, float y, string operationName)
        {
            if (x < Constants.MIN_COORDINATE || x > Constants.MAX_COORDINATE ||
                y < Constants.MIN_COORDINATE || y > Constants.MAX_COORDINATE)
            {
                ComExceptionHandler.LogWarning($"{operationName}: 座標が範囲外 ({x:F1}, {y:F1}) - " +
                    $"有効範囲: {Constants.MIN_COORDINATE}～{Constants.MAX_COORDINATE}pt");
                return false;
            }

            return true;
        }

        #endregion

        #region ユーザーメッセージ

        /// <summary>
        /// 統一された選択エラーメッセージを表示
        /// </summary>
        /// <param name="minimumCount">必要最小図形数</param>
        /// <param name="operationName">操作名</param>
        public static void ShowSelectionError(int minimumCount, string operationName)
        {
            string message = minimumCount == 1
                ? "図形を1つ以上選択してください。"
                : $"{minimumCount}つ以上の図形を選択してください。";

            ComExceptionHandler.LogWarning($"{operationName}: 図形選択不足");
            ShowUserMessage(message, "選択エラー", MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 統一された操作エラーメッセージを表示
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <param name="ex">例外オブジェクト</param>
        public static void ShowOperationError(string operationName, Exception ex)
        {
            string userMessage = ComExceptionHandler.CreateUserErrorMessage(operationName, ex);
            ComExceptionHandler.LogError($"{operationName}エラー", ex);
            ShowUserMessage(userMessage, "エラー", MessageBoxIcon.Error);
        }

        /// <summary>
        /// 統一された成功メッセージを表示
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <param name="details">詳細情報</param>
        /// <param name="showDialog">ダイアログ表示するか</param>
        public static void ShowOperationSuccess(string operationName, string details = null, bool showDialog = false)
        {
            string successMessage = ComExceptionHandler.CreateSuccessMessage(operationName, details);
            ComExceptionHandler.LogDebug($"操作成功: {successMessage}");

            if (showDialog)
            {
                ShowUserMessage(successMessage, "操作完了", MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 統一されたユーザーメッセージ表示
        /// </summary>
        /// <param name="message">メッセージ</param>
        /// <param name="title">タイトル</param>
        /// <param name="icon">アイコン</param>
        private static void ShowUserMessage(string message, string title, MessageBoxIcon icon)
        {
            MessageBox.Show(message, $"Magosa Tools - {title}", MessageBoxButtons.OK, icon);
        }

        /// <summary>
        /// 検証エラーメッセージの作成
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <param name="details">詳細</param>
        /// <returns>フォーマットされたメッセージ</returns>
        private static string CreateValidationMessage(string operationName, string details)
        {
            return $"{operationName}の入力値に問題があります。\n\n{details}";
        }

        #endregion

        #region 図形タイプ検証

        /// <summary>
        /// 四角形図形の検証
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="operationName">操作名</param>
        /// <returns>四角形のみのリストと警告表示結果</returns>
        public static (List<PowerPoint.Shape> rectangles, bool userContinued) ValidateRectangleShapes(
            List<PowerPoint.Shape> shapes, string operationName)
        {
            if (shapes == null || shapes.Count == 0)
            {
                return (new List<PowerPoint.Shape>(), false);
            }

            var rectangles = new List<PowerPoint.Shape>();
            var nonRectangleShapes = new List<string>();

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                if (IsRectangleShape(shape))
                {
                    rectangles.Add(shape);
                }
                else
                {
                    nonRectangleShapes.Add($"図形{i + 1}: {GetShapeTypeName(shape)}");
                }
            }

            // 四角形以外が含まれている場合の警告
            if (nonRectangleShapes.Count > 0)
            {
                var message = $"{operationName}では四角形のみが対象です。\n\n" +
                             "選択中に四角形以外の図形が含まれています：\n" +
                             string.Join("\n", nonRectangleShapes) +
                             "\n\n四角形のみを対象として処理を続行しますか？";

                ComExceptionHandler.LogWarning($"{operationName}: 非四角形図形を検出 ({nonRectangleShapes.Count}個)");

                var result = MessageBox.Show(message, $"Magosa Tools - {operationName}確認",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                return (rectangles, result == DialogResult.Yes);
            }

            return (rectangles, true);
        }

        /// <summary>
        /// 図形が四角形かどうかを判定
        /// </summary>
        public static bool IsRectangleShape(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.HandleComOperation(
                () => shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle ||
                      shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle,
                "図形タイプ判定",
                defaultValue: false,
                throwOnError: false);
        }

        /// <summary>
        /// 図形タイプ名を取得
        /// </summary>
        public static string GetShapeTypeName(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.HandleComOperation(
                () => {
                    switch (shape.Type)
                    {
                        case Microsoft.Office.Core.MsoShapeType.msoAutoShape:
                            return $"オートシェイプ({shape.AutoShapeType})";
                        case Microsoft.Office.Core.MsoShapeType.msoTextBox:
                            return "テキストボックス";
                        case Microsoft.Office.Core.MsoShapeType.msoPicture:
                            return "画像";
                        case Microsoft.Office.Core.MsoShapeType.msoLine:
                            return "線";
                        case Microsoft.Office.Core.MsoShapeType.msoFreeform:
                            return "フリーフォーム";
                        default:
                            return shape.Type.ToString();
                    }
                },
                "図形タイプ名取得",
                defaultValue: "不明な図形",
                throwOnError: false);
        }

        #endregion
    }
}