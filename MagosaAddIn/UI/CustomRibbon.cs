using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI
{
    public partial class CustomRibbon
    {
        private void CustomRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        #region 図形分割機能

        private void btnDivideShape_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selectedShape = GetSelectedShape();
                if (selectedShape != null)
                {
                    ShowDivisionDialog(selectedShape);
                }
                else
                {
                    MessageBox.Show("四角形オブジェクトを選択してください。", "エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private PowerPoint.Shape GetSelectedShape()
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

                    // 四角形またはその他の適切な図形かどうかをチェック
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

        private void ShowDivisionDialog(PowerPoint.Shape shape)
        {
            using (var dialog = new DivisionDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var divider = new ShapeDivider();
                    divider.DivideShape(shape, dialog.Rows, dialog.Columns,
                        dialog.HorizontalMargin, dialog.VerticalMargin);
                }
            }
        }

        #endregion

        #region 基準整列機能

        private void btnAlignToLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToLeft(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を左端揃えしました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("左端揃え", ex.Message);
            }
        }

        private void btnAlignToRight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToRight(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を右端揃えしました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("右端揃え", ex.Message);
            }
        }

        private void btnAlignToTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToTop(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を上端揃えしました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("上端揃え", ex.Message);
            }
        }

        private void btnAlignToBottom_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToBottom(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を下端揃えしました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("下端揃え", ex.Message);
            }
        }

        #endregion

        #region 隣接整列機能

        private void btnAlignLeftToRight_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignLeftToRight(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を左端→右端で整列しました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("左端→右端整列", ex.Message);
            }
        }

        private void btnAlignRightToLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignRightToLeft(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を右端→左端で整列しました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("右端→左端整列", ex.Message);
            }
        }

        private void btnAlignTopToBottom_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignTopToBottom(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を上端→下端で整列しました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("上端→下端整列", ex.Message);
            }
        }

        private void btnAlignBottomToTop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignBottomToTop(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を下端→上端で整列しました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("下端→上端整列", ex.Message);
            }
        }

        #endregion

        #region 拡張整列機能

        private void btnAlignToCenterHorizontal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToCenterHorizontal(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を水平中央揃えしました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("水平中央揃え", ex.Message);
            }
        }

        private void btnAlignToCenterVertical_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    var aligner = new ShapeAligner();
                    aligner.AlignToCenterVertical(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を垂直中央揃えしました。");
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("垂直中央揃え", ex.Message);
            }
        }

        private void btnDistributeHorizontal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 3)
                {
                    var aligner = new ShapeAligner();
                    aligner.DistributeHorizontal(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を水平等間隔配置しました。");
                }
                else
                {
                    MessageBox.Show("3つ以上のオブジェクトを選択してください。", "水平等間隔配置",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("水平等間隔配置", ex.Message);
            }
        }

        private void btnDistributeVertical_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 3)
                {
                    var aligner = new ShapeAligner();
                    aligner.DistributeVertical(shapes);
                    ShowSuccessMessage($"{shapes.Count}個の図形を垂直等間隔配置しました。");
                }
                else
                {
                    MessageBox.Show("3つ以上のオブジェクトを選択してください。", "垂直等間隔配置",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("垂直等間隔配置", ex.Message);
            }
        }

        private void btnArrangeInGrid_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    ShowGridArrangementDialog(shapes);
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("グリッド配置", ex.Message);
            }
        }

        private void btnArrangeInCircle_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var shapes = GetMultipleSelectedShapes();
                if (shapes != null && shapes.Count >= 2)
                {
                    ShowCircleArrangementDialog(shapes);
                }
                else
                {
                    ShowSelectionError();
                }
            }
            catch (Exception ex)
            {
                ShowAlignmentError("円形配置", ex.Message);
            }
        }

        #endregion

        #region ダイアログ表示メソッド

        private void ShowGridArrangementDialog(List<PowerPoint.Shape> shapes)
        {
            using (var dialog = new GridArrangementDialog(shapes.Count))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var aligner = new ShapeAligner();
                    aligner.ArrangeInGrid(shapes, dialog.Columns, dialog.HorizontalSpacing, dialog.VerticalSpacing);
                    ShowSuccessMessage($"{shapes.Count}個の図形を{dialog.Columns}列のグリッドに配置しました。");
                }
            }
        }

        private void ShowCircleArrangementDialog(List<PowerPoint.Shape> shapes)
        {
            using (var dialog = new CircleArrangementDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var aligner = new ShapeAligner();
                    aligner.ArrangeInCircle(shapes, dialog.CenterX, dialog.CenterY, dialog.Radius);
                    ShowSuccessMessage($"{shapes.Count}個の図形を円形に配置しました。");
                }
            }
        }

        #endregion

        #region 共通メソッド

        private List<PowerPoint.Shape> GetMultipleSelectedShapes()
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

        private void ShowSelectionError()
        {
            MessageBox.Show("2つ以上のオブジェクトを選択してください。", "選択エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowAlignmentError(string operation, string message)
        {
            MessageBox.Show($"{operation}エラー: {message}", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowSuccessMessage(string message)
        {
            // 成功メッセージは通常表示しないが、デバッグ時に有効
            System.Diagnostics.Debug.WriteLine($"成功: {message}");

            // 必要に応じてコメントアウトを外す
            // MessageBox.Show(message, "操作完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion
    }

    #region ダイアログクラス

    /// <summary>
    /// グリッド配置用ダイアログ
    /// </summary>
    public partial class GridArrangementDialog : Form
    {
        public int Columns { get; private set; }
        public float HorizontalSpacing { get; private set; }
        public float VerticalSpacing { get; private set; }

        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalSpacing;
        private NumericUpDown numVerticalSpacing;
        private Button btnOK;
        private Button btnCancel;

        public GridArrangementDialog(int shapeCount)
        {
            InitializeComponent(shapeCount);
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Columns = 3;
            HorizontalSpacing = 10.0f;
            VerticalSpacing = 10.0f;
        }

        private void InitializeComponent(int shapeCount)
        {
            this.SuspendLayout();

            this.Text = "グリッド配置設定";
            this.Size = new System.Drawing.Size(300, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 列数設定
            var lblColumns = new Label
            {
                Text = "列数:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(80, 20)
            };

            numColumns = new NumericUpDown
            {
                Location = new System.Drawing.Point(120, 18),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 1,
                Maximum = shapeCount,
                Value = Math.Min(3, shapeCount)
            };

            // 水平間隔設定
            var lblHorizontalSpacing = new Label
            {
                Text = "水平間隔:",
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(80, 20)
            };

            numHorizontalSpacing = new NumericUpDown
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 0,
                Maximum = 200,
                Value = 10,
                DecimalPlaces = 1
            };

            // 垂直間隔設定
            var lblVerticalSpacing = new Label
            {
                Text = "垂直間隔:",
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(80, 20)
            };

            numVerticalSpacing = new NumericUpDown
            {
                Location = new System.Drawing.Point(120, 78),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 0,
                Maximum = 200,
                Value = 10,
                DecimalPlaces = 1
            };

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(70, 120),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new System.Drawing.Point(160, 120),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] {
                lblColumns, numColumns,
                lblHorizontalSpacing, numHorizontalSpacing,
                lblVerticalSpacing, numVerticalSpacing,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Columns = (int)numColumns.Value;
            HorizontalSpacing = (float)numHorizontalSpacing.Value;
            VerticalSpacing = (float)numVerticalSpacing.Value;
        }
    }

    /// <summary>
    /// 円形配置用ダイアログ
    /// </summary>
    public partial class CircleArrangementDialog : Form
    {
        public float CenterX { get; private set; }
        public float CenterY { get; private set; }
        public float Radius { get; private set; }

        private NumericUpDown numCenterX;
        private NumericUpDown numCenterY;
        private NumericUpDown numRadius;
        private Button btnOK;
        private Button btnCancel;
        private Button btnUseCurrentCenter;

        public CircleArrangementDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            CenterX = 400.0f;
            CenterY = 300.0f;
            Radius = 100.0f;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.Text = "円形配置設定";
            this.Size = new System.Drawing.Size(300, 220);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 中心X座標設定
            var lblCenterX = new Label
            {
                Text = "中心X座標:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(80, 20)
            };

            numCenterX = new NumericUpDown
            {
                Location = new System.Drawing.Point(120, 18),
                Size = new System.Drawing.Size(80, 20),
                Minimum = -1000,
                Maximum = 2000,
                Value = 400,
                DecimalPlaces = 1
            };

            // 中心Y座標設定
            var lblCenterY = new Label
            {
                Text = "中心Y座標:",
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(80, 20)
            };

            numCenterY = new NumericUpDown
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new System.Drawing.Size(80, 20),
                Minimum = -1000,
                Maximum = 2000,
                Value = 300,
                DecimalPlaces = 1
            };

            // 半径設定
            var lblRadius = new Label
            {
                Text = "半径:",
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(80, 20)
            };

            numRadius = new NumericUpDown
            {
                Location = new System.Drawing.Point(120, 78),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 10,
                Maximum = 500,
                Value = 100,
                DecimalPlaces = 1
            };

            // 現在の中心を使用ボタン
            btnUseCurrentCenter = new Button
            {
                Text = "選択図形の中心を使用",
                Location = new System.Drawing.Point(20, 110),
                Size = new System.Drawing.Size(150, 25)
            };
            btnUseCurrentCenter.Click += BtnUseCurrentCenter_Click;

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(70, 150),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new System.Drawing.Point(160, 150),
                Size = new System.Drawing.Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] {
                lblCenterX, numCenterX,
                lblCenterY, numCenterY,
                lblRadius, numRadius,
                btnUseCurrentCenter,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void BtnUseCurrentCenter_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActiveWindow?.Selection != null)
                {
                    var selection = app.ActiveWindow.Selection;
                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        var shapes = new List<PowerPoint.Shape>();
                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            shapes.Add(selection.ShapeRange[i]);
                        }

                        if (shapes.Count > 0)
                        {
                            float minLeft = shapes.Min(s => s.Left);
                            float maxRight = shapes.Max(s => s.Left + s.Width);
                            float minTop = shapes.Min(s => s.Top);
                            float maxBottom = shapes.Max(s => s.Top + s.Height);

                            numCenterX.Value = (decimal)((minLeft + maxRight) / 2);
                            numCenterY.Value = (decimal)((minTop + maxBottom) / 2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"中心座標の取得に失敗しました: {ex.Message}", "エラー");
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            CenterX = (float)numCenterX.Value;
            CenterY = (float)numCenterY.Value;
            Radius = (float)numRadius.Value;
        }
    }

    #endregion

}
