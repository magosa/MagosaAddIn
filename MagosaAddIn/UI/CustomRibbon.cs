using MagosaAddIn.Core;
using MagosaAddIn.UI;
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
        private void CustomRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            ComExceptionHandler.LogDebug("Magosa Tools リボンが読み込まれました");
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
            return ComExceptionHandler.HandleComOperation(
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
                throwOnError: false);
        }

        private void ShowDivisionDialog(PowerPoint.Shape shape)
        {
            try
            {
                using (var dialog = new MagosaAddIn.UI.DivisionDialog())
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
                using (var dialog = new MagosaAddIn.UI.GridDivisionDialog(shapes))
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
                using (var dialog = new MagosaAddIn.UI.MarginDialog("水平マージン配置"))
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
                using (var dialog = new MagosaAddIn.UI.MarginDialog("垂直マージン配置"))
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
                using (var dialog = new MagosaAddIn.UI.GridArrangementDialog(shapes.Count))
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
                using (var dialog = new MagosaAddIn.UI.CircleArrangementDialog())
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
            return ComExceptionHandler.HandleComOperation(
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
                throwOnError: false);
        }

        /// <summary>
        /// 図形選択ダイアログを表示
        /// </summary>
        /// <param name="baseShape">基準図形</param>
        private void ShowShapeSelectionDialog(PowerPoint.Shape baseShape)
        {
            try
            {
                using (var dialog = new MagosaAddIn.UI.ShapeSelectionDialog(baseShape))
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
    }
}