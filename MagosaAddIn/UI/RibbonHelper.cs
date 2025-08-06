using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            try
            {
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
            }
            catch (Exception ex)
            {
                throw new Exception($"図形の取得中にエラーが発生しました: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// 複数の図形を取得する
        /// </summary>
        /// <returns>選択された図形のリスト、または null</returns>
        public static List<PowerPoint.Shape> GetMultipleSelectedShapes()
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                throw new Exception($"図形の取得中にエラーが発生しました: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// 選択エラーメッセージを表示する
        /// </summary>
        public static void ShowSelectionError()
        {
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
            MessageBox.Show($"{operation}エラー: {message}", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 成功メッセージを表示する（デバッグ用）
        /// </summary>
        /// <param name="message">成功メッセージ</param>
        public static void ShowSuccessMessage(string message)
        {
            // 成功メッセージは通常表示しないが、デバッグ時に有効
            System.Diagnostics.Debug.WriteLine($"成功: {message}");

            // 必要に応じてコメントアウトを外す
            // MessageBox.Show(message, "操作完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
