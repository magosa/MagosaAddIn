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
            return ComExceptionHandler.ExecuteComOperation(
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
                suppressErrors: true);
        }

        /// <summary>
        /// 複数の図形を取得する
        /// </summary>
        /// <param name="minimumCount">最小要件図形数（デフォルト: 2）</param>
        /// <returns>選択された図形のリスト、または null</returns>
        public static List<PowerPoint.Shape> GetMultipleSelectedShapes(int minimumCount = Constants.MIN_SHAPES_FOR_ALIGNMENT)
        {
            return ComExceptionHandler.ExecuteComOperation(
            () => {
                ComExceptionHandler.LogDebug("GetMultipleSelectedShapes: 開始");

                var app = Globals.ThisAddIn.Application;
                if (app?.ActiveWindow?.Selection == null)
                {
                    ComExceptionHandler.LogDebug("GetMultipleSelectedShapes: app/ActiveWindow/Selectionがnull");
                    return null;
                }

                var selection = app.ActiveWindow.Selection;
                ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: Selection.Type = {selection.Type}");

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: ShapeRange.Count = {selection.ShapeRange.Count}");
                    ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: 最小要件 = {minimumCount}");

                    if (selection.ShapeRange.Count >= minimumCount)
                    {
                        var shapes = new List<PowerPoint.Shape>();
                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            shapes.Add(selection.ShapeRange[i]);
                            ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: 図形{i}を追加 - {selection.ShapeRange[i].Name}");
                        }
                        ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: 成功 - {shapes.Count}個の図形を返す");
                        return shapes;
                    }
                    else
                    {
                        ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: 図形数不足 - {selection.ShapeRange.Count}個 < {minimumCount}個");
                    }
                }
                else
                {
                    ComExceptionHandler.LogDebug($"GetMultipleSelectedShapes: 図形選択ではない - Type = {selection.Type}");
                }

                ComExceptionHandler.LogDebug("GetMultipleSelectedShapes: nullを返す");
                return null;
            },
                "複数図形取得",
                defaultValue: null,
                suppressErrors: true);
        }

        /// <summary>
        /// 選択エラーメッセージを表示する（非推奨 - ErrorHandler.ShowSelectionErrorを使用してください）
        /// </summary>
        [Obsolete("ErrorHandler.ShowSelectionError()を使用してください")]
        public static void ShowSelectionError()
        {
            ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "図形操作");
        }

        /// <summary>
        /// 整列エラーメッセージを表示する（非推奨 - ErrorHandler.ShowOperationErrorを使用してください）
        /// </summary>
        /// <param name="operation">操作名</param>
        /// <param name="message">エラーメッセージ</param>
        [Obsolete("ErrorHandler.ShowOperationError()を使用してください")]
        public static void ShowAlignmentError(string operation, string message)
        {
            ErrorHandler.ShowOperationError(operation, new Exception(message));
        }

        /// <summary>
        /// 整列エラーメッセージを表示する（非推奨 - ErrorHandler.ShowOperationErrorを使用してください）
        /// </summary>
        /// <param name="operation">操作名</param>
        /// <param name="ex">例外オブジェクト</param>
        [Obsolete("ErrorHandler.ShowOperationError()を使用してください")]
        public static void ShowAlignmentError(string operation, Exception ex)
        {
            ErrorHandler.ShowOperationError(operation, ex);
        }

        /// <summary>
        /// 成功メッセージを表示する（非推奨 - ErrorHandler.ShowOperationSuccessを使用してください）
        /// </summary>
        /// <param name="message">成功メッセージ</param>
        [Obsolete("ErrorHandler.ShowOperationSuccess()を使用してください")]
        public static void ShowSuccessMessage(string message)
        {
            ErrorHandler.ShowOperationSuccess("操作", message);
        }

        /// <summary>
        /// PowerPointアプリケーションの状態を確認する
        /// </summary>
        /// <returns>PowerPointが利用可能な場合true</returns>
        public static bool IsPowerPointAvailable()
        {
            return ComExceptionHandler.ExecuteComOperation(
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
                suppressErrors: true);
        }

        /// <summary>
        /// 選択された図形の数を取得する
        /// </summary>
        /// <returns>選択図形数、エラー時は0</returns>
        public static int GetSelectedShapeCount()
        {
            return ComExceptionHandler.ExecuteComOperation(
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
                suppressErrors: true);
        }

        /// <summary>
        /// 図形選択の妥当性をチェックする
        /// </summary>
        /// <param name="minimumCount">必要最小図形数</param>
        /// <returns>選択が妥当な場合true</returns>
        public static bool ValidateShapeSelection(int minimumCount = Constants.MIN_SHAPES_FOR_ALIGNMENT)
        {
            try
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
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("図形選択検証", ex);
                return false;
            }
        }

        /// <summary>
        /// 選択された図形のタイプを分析する
        /// </summary>
        /// <returns>図形タイプの統計情報</returns>
        public static ShapeSelectionInfo AnalyzeSelectedShapes()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var info = new ShapeSelectionInfo();

                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return info;

                    var selection = app.ActiveWindow.Selection;
                    if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                        return info;

                    info.TotalCount = selection.ShapeRange.Count;

                    for (int i = 1; i <= selection.ShapeRange.Count; i++)
                    {
                        var shape = selection.ShapeRange[i];

                        if (ErrorHandler.IsRectangleShape(shape))
                        {
                            info.RectangleCount++;
                        }
                        else
                        {
                            info.NonRectangleCount++;
                            info.NonRectangleTypes.Add(ErrorHandler.GetShapeTypeName(shape));
                        }
                    }

                    return info;
                },
                "図形選択分析",
                defaultValue: new ShapeSelectionInfo(),
                suppressErrors: true);
        }

        /// <summary>
        /// 選択図形の境界を取得する
        /// </summary>
        /// <returns>選択図形の境界情報、エラー時はnull</returns>
        public static ShapeGroupBounds GetSelectedShapesBounds()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var shapes = GetMultipleSelectedShapes();
                    if (shapes == null || shapes.Count == 0)
                        return null;

                    return ShapeGroupBounds.FromShapes(shapes);
                },
                "選択図形境界取得",
                defaultValue: null,
                suppressErrors: true);
        }

        /// <summary>
        /// デバッグ情報を出力する
        /// </summary>
        /// <param name="message">デバッグメッセージ</param>
        public static void LogDebug(string message)
        {
            ComExceptionHandler.LogDebug($"RibbonHelper: {message}");
        }

        /// <summary>
        /// 警告情報を出力する
        /// </summary>
        /// <param name="message">警告メッセージ</param>
        public static void LogWarning(string message)
        {
            ComExceptionHandler.LogWarning($"RibbonHelper: {message}");
        }

        /// <summary>
        /// エラー情報を出力する
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="ex">例外オブジェクト</param>
        public static void LogError(string message, Exception ex = null)
        {
            ComExceptionHandler.LogError($"RibbonHelper: {message}", ex);
        }
    }

    /// <summary>
    /// 図形選択情報を格納するクラス
    /// </summary>
    public class ShapeSelectionInfo
    {
        /// <summary>
        /// 総選択図形数
        /// </summary>
        public int TotalCount { get; set; }

        /// <summary>
        /// 四角形の数
        /// </summary>
        public int RectangleCount { get; set; }

        /// <summary>
        /// 四角形以外の数
        /// </summary>
        public int NonRectangleCount { get; set; }

        /// <summary>
        /// 四角形以外の図形タイプリスト
        /// </summary>
        public List<string> NonRectangleTypes { get; set; } = new List<string>();

        /// <summary>
        /// 四角形のみが選択されているか
        /// </summary>
        public bool IsAllRectangles => NonRectangleCount == 0 && RectangleCount > 0;

        /// <summary>
        /// 選択図形が存在するか
        /// </summary>
        public bool HasShapes => TotalCount > 0;

        /// <summary>
        /// 指定した最小数以上の図形が選択されているか
        /// </summary>
        /// <param name="minimumCount">最小数</param>
        /// <returns>条件を満たす場合true</returns>
        public bool HasMinimumShapes(int minimumCount) => TotalCount >= minimumCount;

        /// <summary>
        /// 選択情報の文字列表現
        /// </summary>
        public override string ToString()
        {
            return $"総数: {TotalCount}, 四角形: {RectangleCount}, その他: {NonRectangleCount}";
        }
    }
}