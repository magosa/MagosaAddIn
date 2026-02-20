using MagosaAddIn.Core;
using MagosaAddIn.UI;
using MagosaAddIn.UI.Dialogs;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI
{
    public partial class CustomRibbon
    {
        private ShapeReplacer shapeReplacer;

        private void CustomRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            ComExceptionHandler.LogDebug("Magosa Tools リボンが読み込まれました");
            shapeReplacer = new ShapeReplacer();
        }

        #region 図形分割機能

        private void btnDivideShape_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selectedShapes = GetSelectedShapesForDivision();
                if (selectedShapes != null && selectedShapes.Count > 0)
                {
                    if (selectedShapes.Count == 1)
                    {
                        // 単一図形の場合（既存機能）
                        ShowDivisionDialog(selectedShapes[0]);
                    }
                    else
                    {
                        // 複数図形の場合（新機能）
                        ShowGridDivisionDialog(selectedShapes);
                    }
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_DIVISION, "図形分割");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("図形分割", ex);
            }
        }

        private List<PowerPoint.Shape> GetSelectedShapesForDivision()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return null;

                    var selection = app.ActiveWindow.Selection;

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                        selection.ShapeRange.Count >= Constants.MIN_SHAPES_FOR_DIVISION)
                    {
                        var allShapes = new List<PowerPoint.Shape>();
                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            allShapes.Add(selection.ShapeRange[i]);
                        }

                        // 四角形図形の検証
                        var (rectangles, userContinued) = ErrorHandler.ValidateRectangleShapes(allShapes, "図形分割");

                        return userContinued ? rectangles : null;
                    }

                    return null;
                },
                "分割対象図形取得",
                defaultValue: null,
                suppressErrors: true);
        }

        private void ShowDivisionDialog(PowerPoint.Shape shape)
        {
            try
            {
                using (var dialog = new DivisionDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var divider = new ShapeDivider();
                        divider.DivideShape(shape, dialog.Rows, dialog.Columns,
                            dialog.HorizontalMargin, dialog.VerticalMargin);

                        ErrorHandler.ShowOperationSuccess("図形分割",
                            $"{dialog.Rows}×{dialog.Columns}グリッドで分割しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("図形分割", ex);
            }
        }

        private void ShowGridDivisionDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new GridDivisionDialog(shapes))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var divider = new ShapeDivider();
                        divider.DivideShapeGroup(shapes, dialog.Rows, dialog.Columns,
                            dialog.HorizontalMargin, dialog.VerticalMargin, dialog.DeleteOriginalShapes);

                        ErrorHandler.ShowOperationSuccess("グリッド分割",
                            $"{shapes.Count}個の図形を{dialog.Rows}×{dialog.Columns}グリッドで分割しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("グリッド分割", ex);
            }
        }

        #endregion

        #region 基準整列機能

        private void btnAlignToLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToLeft(shapes);
                    ErrorHandler.ShowOperationSuccess("左端揃え", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "左端揃え");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("左端揃え", ex);
            }
        }

        private void btnAlignToRight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToRight(shapes);
                    ErrorHandler.ShowOperationSuccess("右端揃え", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "右端揃え");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("右端揃え", ex);
            }
        }

        private void btnAlignToTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToTop(shapes);
                    ErrorHandler.ShowOperationSuccess("上端揃え", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "上端揃え");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("上端揃え", ex);
            }
        }

        private void btnAlignToBottom_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToBottom(shapes);
                    ErrorHandler.ShowOperationSuccess("下端揃え", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "下端揃え");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("下端揃え", ex);
            }
        }

        private void btnAlignToHorizontalCenter_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToHorizontalCenter(shapes);
                    ErrorHandler.ShowOperationSuccess("水平中央揃え", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "水平中央揃え");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("水平中央揃え", ex);
            }
        }

        private void btnAlignToVerticalCenter_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToVerticalCenter(shapes);
                    ErrorHandler.ShowOperationSuccess("垂直中央揃え", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "垂直中央揃え");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("垂直中央揃え", ex);
            }
        }

        #endregion

        #region 隣接整列機能

        private void btnAlignLeftToRight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignLeftToRight(shapes);
                    ErrorHandler.ShowOperationSuccess("左端→右端隣接整列", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "左端→右端整列");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("左端→右端整列", ex);
            }
        }

        private void btnAlignRightToLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignRightToLeft(shapes);
                    ErrorHandler.ShowOperationSuccess("右端→左端隣接整列", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "右端→左端整列");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("右端→左端整列", ex);
            }
        }

        private void btnAlignTopToBottom_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignTopToBottom(shapes);
                    ErrorHandler.ShowOperationSuccess("上端→下端隣接整列", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "上端→下端整列");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("上端→下端整列", ex);
            }
        }

        private void btnAlignBottomToTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignBottomToTop(shapes);
                    ErrorHandler.ShowOperationSuccess("下端→上端隣接整列", $"{shapes.Count}個の図形を整列しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "下端→上端整列");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("下端→上端整列", ex);
            }
        }

        #endregion

        #region 拡張整列機能

        private void btnAlignAndDistributeHorizontal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignAndDistributeHorizontal(shapes);
                    ErrorHandler.ShowOperationSuccess("水平中央揃え・等間隔配置", $"{shapes.Count}個の図形を配置しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "水平中央揃え・等間隔配置");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("水平中央揃え・等間隔配置", ex);
            }
        }

        private void btnAlignAndDistributeVertical_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignAndDistributeVertical(shapes);
                    ErrorHandler.ShowOperationSuccess("垂直中央揃え・等間隔配置", $"{shapes.Count}個の図形を配置しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "垂直中央揃え・等間隔配置");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("垂直中央揃え・等間隔配置", ex);
            }
        }

        private void btnArrangeHorizontalWithMargin_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    ShowHorizontalMarginDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "水平マージン配置");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("水平マージン配置", ex);
            }
        }

        private void btnArrangeVerticalWithMargin_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    ShowVerticalMarginDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "垂直マージン配置");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("垂直マージン配置", ex);
            }
        }

        private void btnArrangeInGrid_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    ShowGridArrangementDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "グリッド配置");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("グリッド配置", ex);
            }
        }

        private void btnArrangeInCircle_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    ShowCircleArrangementDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "円形配置");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("円形配置", ex);
            }
        }

        #endregion

        #region ダイアログ表示メソッド

        private void ShowHorizontalMarginDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new MarginDialog("水平マージン配置"))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var aligner = new ShapeAligner();
                        aligner.ArrangeHorizontalWithMargin(shapes, dialog.Margin);
                        ErrorHandler.ShowOperationSuccess("水平マージン配置",
                            $"{shapes.Count}個の図形をマージン{dialog.Margin:F1}ptで配置しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("水平マージン配置", ex);
            }
        }

        private void ShowVerticalMarginDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new MarginDialog("垂直マージン配置"))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var aligner = new ShapeAligner();
                        aligner.ArrangeVerticalWithMargin(shapes, dialog.Margin);
                        ErrorHandler.ShowOperationSuccess("垂直マージン配置",
                            $"{shapes.Count}個の図形をマージン{dialog.Margin:F1}ptで配置しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("垂直マージン配置", ex);
            }
        }

        private void ShowGridArrangementDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new GridArrangementDialog(shapes.Count))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var aligner = new ShapeAligner();
                        aligner.ArrangeInGrid(shapes, dialog.Columns, dialog.HorizontalSpacing, dialog.VerticalSpacing);
                        ErrorHandler.ShowOperationSuccess("グリッド配置",
                            $"{shapes.Count}個の図形を{dialog.Columns}列のグリッドに配置しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("グリッド配置", ex);
            }
        }

        private void ShowCircleArrangementDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new CircleArrangementDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var aligner = new ShapeAligner();
                        aligner.ArrangeInCircle(shapes, dialog.CenterX, dialog.CenterY, dialog.Radius);
                        ErrorHandler.ShowOperationSuccess("円形配置",
                            $"{shapes.Count}個の図形を円形に配置しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("円形配置", ex);
            }
        }

        #endregion

        #region レイヤー調整機能

        private void btnLayerAdjustment_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_LAYER))
                {
                    ShowLayerAdjustmentDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_LAYER, "レイヤー調整");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("レイヤー調整", ex);
            }
        }

        private void ShowLayerAdjustmentDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new LayerAdjustmentDialog(shapes.Count))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var manager = new ShapeLayerManager();
                        manager.AdjustLayers(shapes, dialog.SelectedOrder);

                        string orderText = GetLayerOrderText(dialog.SelectedOrder);
                        ErrorHandler.ShowOperationSuccess("レイヤー調整",
                            $"{shapes.Count}個の図形の重なり順を調整しました\n方法: {orderText}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("レイヤー調整", ex);
            }
        }

        private string GetLayerOrderText(LayerOrder order)
        {
            switch (order)
            {
                case LayerOrder.SelectionOrderToFront:
                    return "選択順に前面へ配置";
                case LayerOrder.SelectionOrderToBack:
                    return "選択順に背面へ配置";
                case LayerOrder.LeftToRightToFront:
                    return "左から右へ前面に配置";
                case LayerOrder.TopToBottomToFront:
                    return "上から下へ前面に配置";
                default:
                    return "不明";
            }
        }

        #endregion

        #region 自動ナンバリング機能

        private void btnAutoNumbering_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // ★修正: 最小要件1個を指定★
                var shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_NUMBERING);
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_NUMBERING)
                {
                    ShowNumberingDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_NUMBERING, "自動ナンバリング");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("自動ナンバリング", ex);
            }
        }

        private void ShowNumberingDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new NumberingDialog(shapes.Count))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var numbering = new ShapeNumbering();
                        numbering.ApplyNumbering(shapes, dialog.StartNumber, dialog.Increment, 
                            dialog.SelectedFormat, dialog.FontSize);

                        string formatText = GetNumberFormatText(dialog.SelectedFormat);
                        ErrorHandler.ShowOperationSuccess("自動ナンバリング",
                            $"{shapes.Count}個の図形に番号を付けました\n" +
                            $"開始番号: {dialog.StartNumber}, 増分: {dialog.Increment}\n" +
                            $"フォーマット: {formatText}, サイズ: {dialog.FontSize}pt");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("自動ナンバリング", ex);
            }
        }

        private string GetNumberFormatText(NumberFormat format)
        {
            switch (format)
            {
                case NumberFormat.Arabic:
                    return "算用数字 (1, 2, 3...)";
                case NumberFormat.CircledArabic:
                    return "丸数字 (①②③...)";
                case NumberFormat.UpperAlpha:
                    return "大文字アルファベット (A, B, C...)";
                case NumberFormat.LowerAlpha:
                    return "小文字アルファベット (a, b, c...)";
                case NumberFormat.UpperRoman:
                    return "ローマ数字大文字 (I, II, III...)";
                case NumberFormat.LowerRoman:
                    return "ローマ数字小文字 (i, ii, iii...)";
                default:
                    return "不明";
            }
        }

        #endregion

        #region 選択補助機能

        private void btnSelectSameFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 基準となる図形を取得（単一選択が必要）
                var baseShape = GetSingleShapeForFormatSelection();
                if (baseShape == null)
                {
                    ErrorHandler.ShowSelectionError(1, "同一書式選択");
                    return;
                }

                // 選択条件ダイアログを表示
                ShowShapeSelectionDialog(baseShape);
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("同一書式選択", ex);
            }
        }

        /// <summary>
        /// 書式選択用の単一図形を取得
        /// </summary>
        /// <returns>基準図形、または null</returns>
        private PowerPoint.Shape GetSingleShapeForFormatSelection()
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
                        return selection.ShapeRange[1];
                    }

                    return null;
                },
                "書式選択用図形取得",
                defaultValue: null,
                suppressErrors: true);
        }

        /// <summary>
        /// 図形選択ダイアログを表示
        /// </summary>
        /// <param name="baseShape">基準図形</param>
        private void ShowShapeSelectionDialog(PowerPoint.Shape baseShape)
        {
            try
            {
                using (var dialog = new ShapeSelectionDialog(baseShape))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var selector = new ShapeSelector();
                        selector.SelectShapesByFormat(baseShape, dialog.SelectedCriteria);

                        string criteriaText = GetCriteriaDisplayText(dialog.SelectedCriteria);
                        ErrorHandler.ShowOperationSuccess("同一書式選択",
                            $"条件「{criteriaText}」で{dialog.MatchingShapeCount}個の図形を選択しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("同一書式選択", ex);
            }
        }

        /// <summary>
        /// 選択条件の表示用テキストを取得
        /// </summary>
        /// <param name="criteria">選択条件</param>
        /// <returns>表示用テキスト</returns>
        private string GetCriteriaDisplayText(SelectionCriteria criteria)
        {
            switch (criteria)
            {
                case SelectionCriteria.FillColorOnly:
                    return "塗りのカラーコードが同じもの";
                case SelectionCriteria.LineStyleOnly:
                    return "枠線のスタイルが同じもの";
                case SelectionCriteria.FillAndLineStyle:
                    return "塗りと枠線のスタイルが同じもの";
                case SelectionCriteria.ShapeTypeOnly:
                    return "シェイプの種類が同じもの";
                default:
                    return "不明な条件";
            }
        }

        #endregion

        #region ハンドル調整機能

        /// <summary>
        /// 調整ハンドルボタン（一般的な図形の調整ハンドル）
        /// </summary>
        private void btnAdjustmentHandles_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // ★修正: ハンドル調整用の図形取得メソッドを使用★
                var shapes = GetShapesForHandleAdjustment();

                ComExceptionHandler.LogDebug($"btnAdjustmentHandles_Click: 取得した図形数 = {shapes?.Count ?? 0}");

                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT)
                {
                    var adjuster = new ShapeHandleAdjuster();

                    // デバッグ情報出力
                    ComExceptionHandler.LogDebug("=== 調整ハンドル処理開始 ===");
                    adjuster.DebugMultipleShapesInfoLight(shapes);

                    var analysis = adjuster.AnalyzeShapes(shapes);

                    // 分析結果のデバッグ出力
                    ComExceptionHandler.LogDebug($"分析結果: {analysis}");
                    ComExceptionHandler.LogDebug($"HasAdjustmentHandles: {analysis.HasAdjustmentHandles}");
                    ComExceptionHandler.LogDebug($"RecommendedHandleCount: {analysis.RecommendedHandleCount}");

                    if (!analysis.HasAdjustmentHandles)
                    {
                        // エラー前にもう一度詳細確認
                        ComExceptionHandler.LogDebug("=== エラー発生前の最終確認 ===");
                        foreach (var shapeInfo in analysis.ShapeInfos)
                        {
                            ComExceptionHandler.LogDebug($"図形: {shapeInfo.ShapeName}, " +
                                $"タイプ: {shapeInfo.ShapeType}, " +
                                $"ハンドル数: {shapeInfo.AdjustmentCount}, " +
                                $"調整可能: {shapeInfo.IsAdjustmentHandleShape}");
                        }

                        MessageBox.Show(
                            "選択された図形には調整ハンドルがありません。\n\n" +
                            "調整ハンドル対応図形の例:\n" +
                            "・角丸四角形（角丸の調整）\n" +
                            "・吹き出し（尻尾の位置調整）\n" +
                            "・矢印（矢じりの調整）\n" +
                            "・星形（内側の頂点調整）など",
                            "調整ハンドル",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    ShowAdjustmentHandleDialog(shapes, analysis);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "調整ハンドル");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("調整ハンドル", ex);
            }
        }

        /// <summary>
        /// ハンドル調整用の図形取得（1個以上対応）
        /// </summary>
        private List<PowerPoint.Shape> GetShapesForHandleAdjustment()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    ComExceptionHandler.LogDebug("GetShapesForHandleAdjustment: 開始");

                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                    {
                        ComExceptionHandler.LogDebug("GetShapesForHandleAdjustment: app/ActiveWindow/Selectionがnull");
                        return null;
                    }

                    var selection = app.ActiveWindow.Selection;
                    ComExceptionHandler.LogDebug($"GetShapesForHandleAdjustment: Selection.Type = {selection.Type}");

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        ComExceptionHandler.LogDebug($"GetShapesForHandleAdjustment: ShapeRange.Count = {selection.ShapeRange.Count}");

                        // ★修正: 1個以上であればOK★
                        if (selection.ShapeRange.Count >= 1)
                        {
                            var shapes = new List<PowerPoint.Shape>();
                            for (int i = 1; i <= selection.ShapeRange.Count; i++)
                            {
                                shapes.Add(selection.ShapeRange[i]);
                                ComExceptionHandler.LogDebug($"GetShapesForHandleAdjustment: 図形{i}を追加 - {selection.ShapeRange[i].Name}");
                            }
                            ComExceptionHandler.LogDebug($"GetShapesForHandleAdjustment: 成功 - {shapes.Count}個の図形を返す");
                            return shapes;
                        }
                        else
                        {
                            ComExceptionHandler.LogDebug($"GetShapesForHandleAdjustment: 図形が選択されていない");
                        }
                    }
                    else
                    {
                        ComExceptionHandler.LogDebug($"GetShapesForHandleAdjustment: 図形選択ではない - Type = {selection.Type}");
                    }

                    ComExceptionHandler.LogDebug("GetShapesForHandleAdjustment: nullを返す");
                    return null;
                },
                "ハンドル調整用図形取得",
                defaultValue: null,
                suppressErrors: true);
        }

        /// <summary>
        /// 角度ハンドルボタン（角度制御する図形専用）
        /// </summary>
        private void btnAngleHandles_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetShapesForHandleAdjustment();

                ComExceptionHandler.LogDebug($"btnAngleHandles_Click: 取得した図形数 = {shapes?.Count ?? 0}");

                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT)
                {
                    var adjuster = new ShapeHandleAdjuster();

                    // デバッグ情報出力（軽量版）
                    ComExceptionHandler.LogDebug("=== PowerPoint動作調査開始 ===");
                    adjuster.DebugMultipleShapesInfoLight(shapes); // 軽量版を使用

                    var analysis = adjuster.AnalyzeShapes(shapes); // 高速化版を使用

                    // 分析結果のデバッグ出力
                    ComExceptionHandler.LogDebug($"角度ハンドル分析結果: {analysis}");
                    ComExceptionHandler.LogDebug($"HasAngleHandles: {analysis.HasAngleHandles}");
                    ComExceptionHandler.LogDebug($"RecommendedAngleHandleCount: {analysis.RecommendedAngleHandleCount}");

                    if (!analysis.HasAngleHandles)
                    {
                        // エラー前にもう一度詳細確認
                        ComExceptionHandler.LogDebug("=== 角度ハンドルエラー発生前の最終確認 ===");
                        foreach (var shapeInfo in analysis.ShapeInfos)
                        {
                            ComExceptionHandler.LogDebug($"図形: {shapeInfo.ShapeName}, " +
                                $"タイプ: {shapeInfo.ShapeType}, " +
                                $"ハンドル数: {shapeInfo.AdjustmentCount}, " +
                                $"角度ハンドル: {shapeInfo.IsAngleHandleShape}");
                        }

                        MessageBox.Show(
                            "選択された図形には角度ハンドルがありません。\n\n" +
                            "角度ハンドル対応図形:\n" +
                            "・円弧（開始・終了角度）\n" +
                            "・弦（開始・終了角度）\n" +
                            "・扇形/部分円（開始・終了角度）\n" +
                            "・ブロック円弧（角度・内径）\n" +
                            "・ドーナツ（内径比率）\n" +
                            "・三日月（角度調整）など",
                            "角度ハンドル",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    ShowAngleHandleDialog(shapes, analysis);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "角度ハンドル");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("角度ハンドル", ex);
            }
        }


        /// <summary>
        /// 調整リセットボタン
        /// </summary>
        private void btnResetAdjustments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // ★修正: 専用メソッドを使用★
                var shapes = GetShapesForHandleAdjustment();

                ComExceptionHandler.LogDebug($"btnResetAdjustments_Click: 取得した図形数 = {shapes?.Count ?? 0}");

                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT)
                {
                    var adjuster = new ShapeHandleAdjuster();

                    // デバッグ情報出力
                    ComExceptionHandler.LogDebug("=== 調整リセット処理開始 ===");
                    adjuster.DebugMultipleShapesInfoLight(shapes);

                    var analysis = adjuster.AnalyzeShapes(shapes);

                    if (!analysis.HasAdjustmentHandles)
                    {
                        MessageBox.Show(
                            "選択された図形には調整可能なハンドルがありません。",
                            "調整リセット",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    var result = MessageBox.Show(
                        $"選択された{shapes.Count}個の図形の調整をリセットしますか？\n\n" +
                        "・調整ハンドルがデフォルト値（0.5）に戻ります\n" +
                        "・角度ハンドルもデフォルト値に戻ります\n" +
                        $"・対象図形: {analysis.ShapesWithAdjustmentHandles}個",
                        "調整リセット確認",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        adjuster.ResetAdjustments(shapes);
                        ErrorHandler.ShowOperationSuccess("調整リセット",
                            $"{analysis.ShapesWithAdjustmentHandles}個の図形をリセットしました");
                    }
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "調整リセット");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("調整リセット", ex);
            }
        }

        #endregion

        #region ダイアログ表示メソッド

        /// <summary>
        /// 調整ハンドル用ダイアログを表示
        /// </summary>
        private void ShowAdjustmentHandleDialog(List<PowerPoint.Shape> shapes, ShapeHandleAnalysis analysis)
        {
            try
            {
                using (var dialog = new DynamicHandleDialog(shapes, analysis))
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.DialogResult_OK)
                    {
                        var adjuster = new ShapeHandleAdjuster();

                        // mm単位での調整を実行
                        adjuster.AdjustHandlesInMm(shapes, dialog.HandleValues);

                        string handleInfo = string.Join(", ", dialog.HandleValues.Select((v, i) => $"ハンドル{i + 1}={v:F1}mm"));
                        ErrorHandler.ShowOperationSuccess("調整ハンドル設定",
                            $"{shapes.Count}個の図形を調整しました\n設定値: {handleInfo}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("調整ハンドル設定", ex);
            }
        }

        /// <summary>
        /// 角度ハンドル用ダイアログを表示
        /// </summary>
        private void ShowAngleHandleDialog(List<PowerPoint.Shape> shapes, ShapeHandleAnalysis analysis)
        {
            try
            {
                using (var dialog = new DynamicAngleHandleDialog(shapes, analysis))
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.DialogResult_OK)
                    {
                        var adjuster = new ShapeHandleAdjuster();

                        // 度数単位での角度調整を実行
                        adjuster.AdjustAngleHandlesInDegree(shapes, dialog.HandleValues);

                        // 角度ハンドル図形の種類を取得
                        var angleShapeTypes = analysis.ShapeInfos
                            .Where(info => info.IsAngleHandleShape)
                            .Select(info => info.GetDisplayShapeType())
                            .Distinct()
                            .ToList();

                        string shapeTypeInfo = string.Join(", ", angleShapeTypes);
                        string handleInfo = string.Join(", ", dialog.HandleValues.Select((v, i) => $"角度{i + 1}={v:F1}°"));

                        ErrorHandler.ShowOperationSuccess("角度ハンドル設定",
                            $"{analysis.ShapesWithAngleHandles}個の角度ハンドル図形を調整しました\n" +
                            $"図形タイプ: {shapeTypeInfo}\n設定値: {handleInfo}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("角度ハンドル設定", ex);
            }
        }

        #endregion

        #region サイズ調整機能

        /// <summary>
        /// 基準サイズ適用ボタン
        /// </summary>
        private void btnResizeToReference_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var resizer = new ShapeResizer();
                    resizer.ResizeToReference(shapes, ResizeMode.KeepCenter);
                    ErrorHandler.ShowOperationSuccess("基準サイズ適用", 
                        $"{shapes.Count}個の図形を基準図形のサイズに調整しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "基準サイズ適用");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("基準サイズ適用", ex);
            }
        }

        /// <summary>
        /// 幅統一（比率保持）ボタン
        /// </summary>
        private void btnResizeToWidthKeepRatio_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var resizer = new ShapeResizer();
                    resizer.ResizeToWidthKeepRatio(shapes);
                    ErrorHandler.ShowOperationSuccess("幅統一・比率保持", 
                        $"{shapes.Count}個の図形の幅を統一しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "幅統一・比率保持");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("幅統一・比率保持", ex);
            }
        }

        /// <summary>
        /// 高さ統一（比率保持）ボタン
        /// </summary>
        private void btnResizeToHeightKeepRatio_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (RibbonHelper.ValidateShapeSelection(Constants.MIN_SHAPES_FOR_ALIGNMENT))
                {
                    var resizer = new ShapeResizer();
                    resizer.ResizeToHeightKeepRatio(shapes);
                    ErrorHandler.ShowOperationSuccess("高さ統一・比率保持", 
                        $"{shapes.Count}個の図形の高さを統一しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ALIGNMENT, "高さ統一・比率保持");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("高さ統一・比率保持", ex);
            }
        }

        /// <summary>
        /// 最大サイズ統一ボタン
        /// </summary>
        private void btnResizeToMaximum_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_RESIZE);
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_RESIZE)
                {
                    var resizer = new ShapeResizer();
                    resizer.ResizeToMaximum(shapes, ResizeMode.KeepCenter);
                    ErrorHandler.ShowOperationSuccess("最大サイズ統一", 
                        $"{shapes.Count}個の図形を最大サイズに統一しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_RESIZE, "最大サイズ統一");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("最大サイズ統一", ex);
            }
        }

        /// <summary>
        /// 最小サイズ統一ボタン
        /// </summary>
        private void btnResizeToMinimum_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_RESIZE);
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_RESIZE)
                {
                    var resizer = new ShapeResizer();
                    resizer.ResizeToMinimum(shapes, ResizeMode.KeepCenter);
                    ErrorHandler.ShowOperationSuccess("最小サイズ統一", 
                        $"{shapes.Count}個の図形を最小サイズに統一しました");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_RESIZE, "最小サイズ統一");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("最小サイズ統一", ex);
            }
        }

        /// <summary>
        /// 拡大縮小・固定サイズボタン（ダイアログ表示）
        /// </summary>
        private void btnResizeDialog_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_RESIZE);
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_RESIZE)
                {
                    ShowResizeDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_RESIZE, "サイズ調整");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("サイズ調整", ex);
            }
        }

        private void ShowResizeDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new ShapeResizeDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var resizer = new ShapeResizer();

                        if (dialog.UsePercentage)
                        {
                            resizer.ResizeByPercentage(shapes, dialog.Percentage);
                            ErrorHandler.ShowOperationSuccess("パーセント拡大縮小",
                                $"{shapes.Count}個の図形を{dialog.Percentage:F1}%に調整しました");
                        }
                        else
                        {
                            resizer.ResizeToFixedSize(shapes, dialog.Width, dialog.Height, 
                                dialog.Unit, dialog.KeepRatio, ResizeMode.KeepCenter);
                            string unitText = dialog.Unit == SizeUnit.Point ? "pt" : 
                                            dialog.Unit == SizeUnit.Millimeter ? "mm" : "cm";
                            ErrorHandler.ShowOperationSuccess("固定サイズ設定",
                                $"{shapes.Count}個の図形を{dialog.Width:F1}×{dialog.Height:F1}{unitText}に調整しました");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("サイズ調整", ex);
            }
        }

        #endregion

        #region 配列複製機能

        /// <summary>
        /// 円形配列ボタン
        /// </summary>
        private void btnCircularArray_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_ARRAY);
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_ARRAY)
                {
                    ShowCircularArrayDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ARRAY, "円形配列");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("円形配列", ex);
            }
        }

        private void ShowCircularArrayDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new CircularArrayDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var arrayer = new ShapeArrayer();
                        arrayer.CircularArray(shapes, dialog.Options);
                        ErrorHandler.ShowOperationSuccess("円形配列",
                            $"{shapes.Count}個の図形を{dialog.Options.Count}個に円形配列しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("円形配列", ex);
            }
        }

        /// <summary>
        /// グリッド配列ボタン
        /// </summary>
        private void btnGridArray_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_ARRAY);
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_ARRAY)
                {
                    ShowGridArrayDialog(shapes);
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_ARRAY, "グリッド配列");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("グリッド配列", ex);
            }
        }

        private void ShowGridArrayDialog(List<PowerPoint.Shape> shapes)
        {
            try
            {
                using (var dialog = new GridArrayDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var arrayer = new ShapeArrayer();
                        arrayer.GridArray(shapes, dialog.Options);
                        int totalShapes = dialog.Options.Rows * dialog.Options.Columns;
                        ErrorHandler.ShowOperationSuccess("グリッド配列",
                            $"{shapes.Count}個の図形を{dialog.Options.Rows}×{dialog.Options.Columns}グリッドに配列しました");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("グリッド配列", ex);
            }
        }


        #endregion

        #region 図形置き換え機能

        /// <summary>
        /// 選択完了ボタン（置き換え対象図形を記憶）
        /// </summary>
        private void btnSaveShapes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= Constants.MIN_SHAPES_FOR_REPLACEMENT)
                {
                    shapeReplacer.SaveShapes(shapes);
                    UpdateSavedCountLabel();
                    ErrorHandler.ShowOperationSuccess("図形記憶", 
                        $"{shapes.Count}個の図形を記憶しました。\n次にテンプレート図形を1つ選択して「置き換え実行」をクリックしてください。");
                }
                else
                {
                    ErrorHandler.ShowSelectionError(Constants.MIN_SHAPES_FOR_REPLACEMENT, "図形記憶");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("図形記憶", ex);
            }
        }

        /// <summary>
        /// 置き換え実行ボタン
        /// </summary>
        private void btnReplaceShapes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 記憶図形チェック
                int savedCount = shapeReplacer.GetSavedShapeCount();
                if (savedCount == 0)
                {
                    MessageBox.Show(
                        "先に置き換え対象の図形を選択して「選択完了」してください。",
                        "図形置き換え",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // テンプレート図形チェック（1個のみ）
                var templateShape = GetSingleShapeForReplacement();
                if (templateShape == null)
                {
                    MessageBox.Show(
                        "テンプレート図形を1つ選択してください。",
                        "図形置き換え",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // ダイアログ表示
                ShowReplacementDialog(templateShape, savedCount);
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("図形置き換え", ex);
            }
        }

        /// <summary>
        /// 置き換え用の単一図形を取得（テンプレート図形）
        /// </summary>
        private PowerPoint.Shape GetSingleShapeForReplacement()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow?.Selection == null)
                        return null;

                    var selection = app.ActiveWindow.Selection;

                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                        selection.ShapeRange.Count == Constants.TEMPLATE_SHAPE_COUNT)
                    {
                        return selection.ShapeRange[1];
                    }

                    return null;
                },
                "テンプレート図形取得",
                defaultValue: null,
                suppressErrors: true);
        }

        /// <summary>
        /// 図形置き換えダイアログを表示
        /// </summary>
        private void ShowReplacementDialog(PowerPoint.Shape templateShape, int savedShapeCount)
        {
            try
            {
                string templateShapeName = ComExceptionHandler.ExecuteComOperation(
                    () => templateShape.Name,
                    "テンプレート図形名取得",
                    defaultValue: "不明な図形",
                    suppressErrors: true);

                using (var dialog = new ShapeReplacementDialog(savedShapeCount, templateShapeName))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var options = new ReplacementOptions
                        {
                            SizeMode = dialog.SelectedSizeMode,
                            InheritStyle = dialog.InheritStyle,
                            InheritText = dialog.InheritText
                        };

                        shapeReplacer.ReplaceShapes(templateShape, options);
                        UpdateSavedCountLabel();

                        string sizeInfo = dialog.SelectedSizeMode == SizeMode.KeepOriginal ? "元のサイズ" : "テンプレートサイズ";
                        string styleInfo = dialog.InheritStyle ? "スタイル継承あり" : "テンプレートスタイル使用";
                        string textInfo = dialog.InheritText ? "テキスト継承あり" : "テンプレートテキスト使用";

                        ErrorHandler.ShowOperationSuccess("図形置き換え",
                            $"{savedShapeCount}個の図形を置き換えました。\n" +
                            $"サイズ: {sizeInfo}\n" +
                            $"{styleInfo}, {textInfo}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("図形置き換え", ex);
            }
        }

        /// <summary>
        /// 記憶図形数ラベルを更新
        /// </summary>
        private void UpdateSavedCountLabel()
        {
            int count = shapeReplacer.GetSavedShapeCount();
            lblSavedCount.Label = $"記憶: {count}個";
        }

        #endregion

        #region テーマカラー生成機能

        /// <summary>
        /// テーマカラー生成ボタン
        /// </summary>
        private void btnThemeColorGenerator_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShowThemeColorDialog();
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("テーマカラー生成", ex);
            }
        }

        /// <summary>
        /// テーマカラー生成ダイアログを表示
        /// </summary>
        private void ShowThemeColorDialog()
        {
            try
            {
                using (var dialog = new ThemeColorDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var generator = new ThemeColorGenerator();
                        var arranger = new ColorPaletteArranger();

                        // 選択図形を取得（オプション）
                        List<PowerPoint.Shape> shapes = null;
                        if (dialog.ApplyToShapes)
                        {
                            shapes = RibbonHelper.GetMultipleSelectedShapes(Constants.MIN_SHAPES_FOR_THEME_COLOR);
                        }

                        // 色を適用
                        if (dialog.ApplyToShapes && shapes != null && shapes.Count > 0)
                        {
                            arranger.ApplyColorsToShapes(shapes, dialog.GeneratedColors, ColorApplyMode.Sequential);
                        }

                        // パレット配置
                        if (dialog.ArrangePalette)
                        {
                            if (dialog.LightnessSteps > 1)
                            {
                                // 明度バリエーション付きでグリッド配置
                                var colorMatrix = generator.GenerateLightnessVariations(
                                    dialog.GeneratedColors, dialog.LightnessSteps);
                                arranger.ArrangeColorGrid(colorMatrix);
                            }
                            else
                            {
                                // 単一行で配置
                                arranger.ArrangeColorRow(dialog.GeneratedColors);
                            }
                        }

                        // 成功メッセージ
                        string schemeName = ThemeColorGenerator.GetSchemeDisplayName(dialog.SelectedScheme);
                        string message = $"テーマカラーを生成しました\n" +
                                       $"配色パターン: {schemeName}\n" +
                                       $"色数: {dialog.GeneratedColors.Count}色";

                        if (dialog.ApplyToShapes && shapes != null && shapes.Count > 0)
                        {
                            message += $"\n{shapes.Count}個の図形に適用しました";
                        }

                        if (dialog.ArrangePalette)
                        {
                            message += "\nカラーパレットを配置しました";
                        }

                        ErrorHandler.ShowOperationSuccess("テーマカラー生成", message);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("テーマカラー生成", ex);
            }
        }

        #endregion
    }
}
