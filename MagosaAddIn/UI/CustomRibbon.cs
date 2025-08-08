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
                if (selectedShapes != null)
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
                    MessageBox.Show("四角形オブジェクトを1つ以上選択してください。", "図形分割",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("図形分割", ex);
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
                        selection.ShapeRange.Count >= 1)
                    {
                        var shapes = new List<PowerPoint.Shape>();
                        var nonRectangleShapes = new List<string>();

                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            var shape = selection.ShapeRange[i];

                            // 四角形かどうかをチェック
                            if (shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle ||
                                shape.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle)
                            {
                                shapes.Add(shape);
                            }
                            else
                            {
                                nonRectangleShapes.Add($"図形{i}: {GetShapeTypeName(shape)}");
                            }
                        }

                        // 四角形以外の図形が含まれている場合の警告
                        if (nonRectangleShapes.Count > 0)
                        {
                            var message = "選択中に四角形以外の図形が含まれています：\n\n" +
                                         string.Join("\n", nonRectangleShapes) +
                                         "\n\n四角形のみを対象として処理を続行しますか？";

                            var result = MessageBox.Show(message, "図形分割 - 確認",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                            {
                                return null;
                            }
                        }

                        return shapes.Count > 0 ? shapes : null;
                    }

                    return null;
                },
                "分割対象図形取得",
                defaultValue: null,
                throwOnError: false);
        }

        private string GetShapeTypeName(PowerPoint.Shape shape)
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

                        string successMessage = ComExceptionHandler.CreateSuccessMessage(
                            "図形分割",
                            $"{dialog.Rows}×{dialog.Columns}グリッドで分割しました");
                        RibbonHelper.ShowSuccessMessage(successMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("図形分割", ex);
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

                        string successMessage = ComExceptionHandler.CreateSuccessMessage(
                            "グリッド分割",
                            $"{shapes.Count}個の図形を{dialog.Rows}×{dialog.Columns}グリッドで分割しました");
                        RibbonHelper.ShowSuccessMessage(successMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("グリッド分割", ex);
            }
        }

        #endregion

        #region 基準整列機能

        private void btnAlignToLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToLeft(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "左端揃え", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("左端揃え", ex);
            }
        }

        private void btnAlignToRight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToRight(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "右端揃え", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("右端揃え", ex);
            }
        }

        private void btnAlignToTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToTop(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "上端揃え", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("上端揃え", ex);
            }
        }

        private void btnAlignToBottom_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToBottom(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "下端揃え", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("下端揃え", ex);
            }
        }

        #endregion

        #region 隣接整列機能

        private void btnAlignLeftToRight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignLeftToRight(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "左端→右端隣接整列", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("左端→右端整列", ex);
            }
        }

        private void btnAlignRightToLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignRightToLeft(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "右端→左端隣接整列", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("右端→左端整列", ex);
            }
        }

        private void btnAlignTopToBottom_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignTopToBottom(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "上端→下端隣接整列", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("上端→下端整列", ex);
            }
        }

        private void btnAlignBottomToTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignBottomToTop(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "下端→上端隣接整列", $"{shapes.Count}個の図形を整列しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("下端→上端整列", ex);
            }
        }

        #endregion

        #region 拡張整列機能

        private void btnAlignAndDistributeHorizontal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignAndDistributeHorizontal(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "水平中央揃え・等間隔配置", $"{shapes.Count}個の図形を配置しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("水平中央揃え・等間隔配置", ex);
            }
        }

        private void btnAlignAndDistributeVertical_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignAndDistributeVertical(shapes);
                    string successMessage = ComExceptionHandler.CreateSuccessMessage(
                        "垂直中央揃え・等間隔配置", $"{shapes.Count}個の図形を配置しました");
                    RibbonHelper.ShowSuccessMessage(successMessage);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("垂直中央揃え・等間隔配置", ex);
            }
        }

        private void btnArrangeHorizontalWithMargin_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    ShowHorizontalMarginDialog(shapes);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("水平マージン配置", ex);
            }
        }

        private void btnArrangeVerticalWithMargin_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    ShowVerticalMarginDialog(shapes);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("垂直マージン配置", ex);
            }
        }

        private void btnArrangeInGrid_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    ShowGridArrangementDialog(shapes);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("グリッド配置", ex);
            }
        }

        private void btnArrangeInCircle_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = RibbonHelper.GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    ShowCircleArrangementDialog(shapes);
                }
                else
                {
                    RibbonHelper.ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("円形配置", ex);
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
                        string successMessage = ComExceptionHandler.CreateSuccessMessage(
                            "水平マージン配置",
                            $"{shapes.Count}個の図形をマージン{dialog.Margin:F1}ptで配置しました");
                        RibbonHelper.ShowSuccessMessage(successMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("水平マージン配置", ex);
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
                        string successMessage = ComExceptionHandler.CreateSuccessMessage(
                            "垂直マージン配置",
                            $"{shapes.Count}個の図形をマージン{dialog.Margin:F1}ptで配置しました");
                        RibbonHelper.ShowSuccessMessage(successMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("垂直マージン配置", ex);
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
                        string successMessage = ComExceptionHandler.CreateSuccessMessage(
                            "グリッド配置",
                            $"{shapes.Count}個の図形を{dialog.Columns}列のグリッドに配置しました");
                        RibbonHelper.ShowSuccessMessage(successMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("グリッド配置", ex);
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
                        string successMessage = ComExceptionHandler.CreateSuccessMessage(
                            "円形配置",
                            $"{shapes.Count}個の図形を円形に配置しました");
                        RibbonHelper.ShowSuccessMessage(successMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                RibbonHelper.ShowAlignmentError("円形配置", ex);
            }
        }

        #endregion
    }
}