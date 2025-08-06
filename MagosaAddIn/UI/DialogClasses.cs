using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI
{
    /// <summary>
    /// 図形分割設定用ダイアログ
    /// </summary>
    public partial class DivisionDialog : Form
    {
        public int Rows { get; private set; }
        public int Columns { get; private set; }
        public float HorizontalMargin { get; private set; }
        public float VerticalMargin { get; private set; }

        private NumericUpDown numRows;
        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalMargin;
        private NumericUpDown numVerticalMargin;
        private Button btnOK;
        private Button btnCancel;
        private CheckBox chkLinkMargins;

        public DivisionDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Rows = 2;
            Columns = 2;
            HorizontalMargin = 2.0f;
            VerticalMargin = 2.0f;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "図形分割設定";
            this.Size = new Size(350, 280);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 行数設定
            var lblRows = new Label
            {
                Text = "行数:",
                Location = new Point(20, 20),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numRows = new NumericUpDown
            {
                Location = new Point(120, 18),
                Size = new Size(80, 20),
                Minimum = 1,
                Maximum = 50,
                Value = 2
            };

            // 列数設定
            var lblColumns = new Label
            {
                Text = "列数:",
                Location = new Point(20, 50),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numColumns = new NumericUpDown
            {
                Location = new Point(120, 48),
                Size = new Size(80, 20),
                Minimum = 1,
                Maximum = 50,
                Value = 2
            };

            // 水平マージン設定
            var lblHorizontalMargin = new Label
            {
                Text = "水平マージン:",
                Location = new Point(20, 80),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numHorizontalMargin = new NumericUpDown
            {
                Location = new Point(120, 78),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 100,
                Value = 2,
                DecimalPlaces = 1,
                Increment = 0.5m
            };

            var lblHorizontalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 80),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 垂直マージン設定
            var lblVerticalMargin = new Label
            {
                Text = "垂直マージン:",
                Location = new Point(20, 110),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numVerticalMargin = new NumericUpDown
            {
                Location = new Point(120, 108),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 100,
                Value = 2,
                DecimalPlaces = 1,
                Increment = 0.5m
            };

            var lblVerticalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 110),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            //// マージン連動チェックボックス
            //chkLinkMargins = new CheckBox
            //{
            //    Text = "水平・垂直マージンを連動",
            //    Location = new Point(20, 140),
            //    Size = new Size(200, 20),
            //    Checked = true
            //};
            //chkLinkMargins.CheckedChanged += ChkLinkMargins_CheckedChanged;

            // プレビューラベル
            var lblPreview = new Label
            {
                Text = "プレビュー: 2×2 グリッド",
                Location = new Point(20, 170),
                Size = new Size(200, 20),
                ForeColor = Color.Gray
            };

            // 値変更時のプレビュー更新
            numRows.ValueChanged += (s, e) => UpdatePreview(lblPreview);
            numColumns.ValueChanged += (s, e) => UpdatePreview(lblPreview);

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(120, 210),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(210, 210),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblRows, numRows,
                lblColumns, numColumns,
                lblHorizontalMargin, numHorizontalMargin, lblHorizontalUnit,
                lblVerticalMargin, numVerticalMargin, lblVerticalUnit,
                chkLinkMargins,
                lblPreview,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void ChkLinkMargins_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLinkMargins.Checked)
            {
                numVerticalMargin.Value = numHorizontalMargin.Value;
                numVerticalMargin.Enabled = false;
                numHorizontalMargin.ValueChanged += SyncMargins;
            }
            else
            {
                numVerticalMargin.Enabled = true;
                numHorizontalMargin.ValueChanged -= SyncMargins;
            }
        }

        private void SyncMargins(object sender, EventArgs e)
        {
            if (chkLinkMargins.Checked)
            {
                numVerticalMargin.Value = numHorizontalMargin.Value;
            }
        }

        private void UpdatePreview(Label lblPreview)
        {
            lblPreview.Text = $"プレビュー: {numRows.Value}×{numColumns.Value} グリッド";
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Rows = (int)numRows.Value;
            Columns = (int)numColumns.Value;
            HorizontalMargin = (float)numHorizontalMargin.Value;
            VerticalMargin = (float)numVerticalMargin.Value;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numRows?.Dispose();
                numColumns?.Dispose();
                numHorizontalMargin?.Dispose();
                numVerticalMargin?.Dispose();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                chkLinkMargins?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// マージン設定用ダイアログ
    /// </summary>
    public partial class MarginDialog : Form
    {
        public float Margin { get; private set; }

        private NumericUpDown numMargin;
        private Button btnOK;
        private Button btnCancel;

        public MarginDialog(string title)
        {
            InitializeComponent(title);
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Margin = 10.0f;
        }

        private void InitializeComponent(string title)
        {
            this.SuspendLayout();

            this.Text = title;
            this.Size = new Size(280, 150);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // マージン設定
            var lblMargin = new Label
            {
                Text = "マージン:",
                Location = new Point(20, 20),
                Size = new Size(60, 20)
            };

            numMargin = new NumericUpDown
            {
                Location = new Point(90, 18),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 200,
                Value = 10,
                DecimalPlaces = 1,
                Increment = 0.5m
            };

            var lblUnit = new Label
            {
                Text = "pt",
                Location = new Point(180, 20),
                Size = new Size(20, 20)
            };

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(70, 70),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(160, 70),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] {
                lblMargin, numMargin, lblUnit,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Margin = (float)numMargin.Value;
        }
    }

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
            this.Size = new Size(300, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 列数設定
            var lblColumns = new Label
            {
                Text = "列数:",
                Location = new Point(20, 20),
                Size = new Size(80, 20)
            };

            numColumns = new NumericUpDown
            {
                Location = new Point(120, 18),
                Size = new Size(80, 20),
                Minimum = 1,
                Maximum = shapeCount,
                Value = Math.Min(3, shapeCount)
            };

            // 水平間隔設定
            var lblHorizontalSpacing = new Label
            {
                Text = "水平間隔:",
                Location = new Point(20, 50),
                Size = new Size(80, 20)
            };

            numHorizontalSpacing = new NumericUpDown
            {
                Location = new Point(120, 48),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 200,
                Value = 10,
                DecimalPlaces = 1
            };

            // 垂直間隔設定
            var lblVerticalSpacing = new Label
            {
                Text = "垂直間隔:",
                Location = new Point(20, 80),
                Size = new Size(80, 20)
            };

            numVerticalSpacing = new NumericUpDown
            {
                Location = new Point(120, 78),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 200,
                Value = 10,
                DecimalPlaces = 1
            };

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(70, 120),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(160, 120),
                Size = new Size(75, 25),
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
            this.Size = new Size(300, 220);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 中心X座標設定
            var lblCenterX = new Label
            {
                Text = "中心X座標:",
                Location = new Point(20, 20),
                Size = new Size(80, 20)
            };

            numCenterX = new NumericUpDown
            {
                Location = new Point(120, 18),
                Size = new Size(80, 20),
                Minimum = -1000,
                Maximum = 2000,
                Value = 400,
                DecimalPlaces = 1
            };

            // 中心Y座標設定
            var lblCenterY = new Label
            {
                Text = "中心Y座標:",
                Location = new Point(20, 50),
                Size = new Size(80, 20)
            };

            numCenterY = new NumericUpDown
            {
                Location = new Point(120, 48),
                Size = new Size(80, 20),
                Minimum = -1000,
                Maximum = 2000,
                Value = 300,
                DecimalPlaces = 1
            };

            // 半径設定
            var lblRadius = new Label
            {
                Text = "半径:",
                Location = new Point(20, 80),
                Size = new Size(80, 20)
            };

            numRadius = new NumericUpDown
            {
                Location = new Point(120, 78),
                Size = new Size(80, 20),
                Minimum = 10,
                Maximum = 500,
                Value = 100,
                DecimalPlaces = 1
            };

            // 現在の中心を使用ボタン
            btnUseCurrentCenter = new Button
            {
                Text = "選択図形の中心を使用",
                Location = new Point(20, 110),
                Size = new Size(150, 25)
            };
            btnUseCurrentCenter.Click += BtnUseCurrentCenter_Click;

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(70, 150),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(160, 150),
                Size = new Size(75, 25),
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
}