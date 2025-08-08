using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI
{
    /// <summary>
    /// リボン機能の共通メソッドを提供するヘルパークラス
    /// </summary>
    public static class RibbonHelper
    {
        /// <summary>
        /// 単一の図形を取得する
        /// </summary>
        /// <returns>選択された図形、または null</returns>
        public static PowerPoint.Shape GetSingleSelectedShape()
        {
            return ComExceptionHandler.HandleComOperation(
                () => {
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return null;

                    var selection = app.ActiveWindow.Selection;

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                        selection.ShapeRange.Count == 1)
                    {
                        var shape = selection.ShapeRange[1];

                        if (shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle ||
                            shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape)
                        {
                            return shape;
                        }
                    }

                    return null;
                },
                "単一図形取得",
                defaultValue: null,
                throwOnError: false);
        }

        /// <summary>
        /// 複数の図形を取得する
        /// </summary>
        /// <returns>選択された図形のリスト、または null</returns>
        public static List<PowerPoint.Shape> GetMultipleSelectedShapes()
        {
            return ComExceptionHandler.HandleComOperation(
                () => {
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return null;

                    var selection = app.ActiveWindow.Selection;

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                        selection.ShapeRange.Count >= 2)
                    {
                        var shapes = new List<PowerPoint.Shape>();
                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            shapes.Add(selection.ShapeRange[i]);
                        }
                        return shapes;
                    }

                    return null;
                },
                "複数図形取得",
                defaultValue: null,
                throwOnError: false);
        }

        /// <summary>
        /// 選択エラーメッセージを表示する
        /// </summary>
        public static void ShowSelectionError()
        {
            ComExceptionHandler.LogWarning("図形選択不足: 2つ以上のオブジェクトが必要");
            MessageBox.Show("2つ以上のオブジェクトを選択してください。", "選択エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 整列エラーメッセージを表示する
        /// </summary>
        /// <param name="operation">操作名</param>
        /// <param name="message">エラーメッセージ</param>
        public static void ShowAlignmentError(string operation, string message)
        {
            string errorMessage = ComExceptionHandler.CreateUserErrorMessage(operation, new Exception(message));
            ComExceptionHandler.LogError($"{operation}エラー", new Exception(message));
            MessageBox.Show(errorMessage, "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 整列エラーメッセージを表示する（例外オブジェクト版）
        /// </summary>
        /// <param name="operation">操作名</param>
        /// <param name="ex">例外オブジェクト</param>
        public static void ShowAlignmentError(string operation, Exception ex)
        {
            string errorMessage = ComExceptionHandler.CreateUserErrorMessage(operation, ex);
            ComExceptionHandler.LogError($"{operation}エラー", ex);
            MessageBox.Show(errorMessage, "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 成功メッセージを表示する（デバッグ用）
        /// </summary>
        /// <param name="message">成功メッセージ</param>
        public static void ShowSuccessMessage(string message)
        {
            // 成功メッセージは統一されたログ出力を使用
            ComExceptionHandler.LogDebug($"操作成功: {message}");

            // 必要に応じてコメントアウトを外す
            // string successMessage = ComExceptionHandler.CreateSuccessMessage("操作", message);
            // MessageBox.Show(successMessage, "操作完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// PowerPointアプリケーションの状態を確認する
        /// </summary>
        /// <returns>PowerPointが利用可能な場合true</returns>
        public static bool IsPowerPointAvailable()
        {
            return ComExceptionHandler.HandleComOperation(
                () => {
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return false;

                    // アクティブウィンドウの存在確認
                    var testWindow = app.ActiveWindow;
                    return testWindow != null;
                },
                "PowerPoint状態確認",
                defaultValue: false,
                throwOnError: false);
        }

        /// <summary>
        /// 選択された図形の数を取得する
        /// </summary>
        /// <returns>選択図形数、エラー時は0</returns>
        public static int GetSelectedShapeCount()
        {
            return ComExceptionHandler.HandleComOperation(
                () => {
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return 0;

                    var selection = app.ActiveWindow.Selection;

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        return selection.ShapeRange.Count;
                    }

                    return 0;
                },
                "選択図形数取得",
                defaultValue: 0,
                throwOnError: false);
        }

        /// <summary>
        /// 図形選択の妥当性をチェックする
        /// </summary>
        /// <param name="minimumCount">必要最小図形数</param>
        /// <returns>選択が妥当な場合true</returns>
        public static bool ValidateShapeSelection(int minimumCount = 2)
        {
            if (!IsPowerPointAvailable())
            {
                ComExceptionHandler.LogWarning("PowerPointが利用できません");
                return false;
            }

            int selectedCount = GetSelectedShapeCount();
            bool isValid = selectedCount >= minimumCount;

            if (!isValid)
            {
                ComExceptionHandler.LogWarning($"図形選択不足: {selectedCount}個選択済み、{minimumCount}個必要");
            }
            else
            {
                ComExceptionHandler.LogDebug($"図形選択確認: {selectedCount}個選択済み");
            }

            return isValid;
        }
    }
}