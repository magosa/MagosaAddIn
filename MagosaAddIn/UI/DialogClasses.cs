using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;

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
            Rows = Constants.DEFAULT_ROWS;
            Columns = Constants.DEFAULT_COLUMNS;
            HorizontalMargin = Constants.DEFAULT_HORIZONTAL_MARGIN;
            VerticalMargin = Constants.DEFAULT_VERTICAL_MARGIN;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "グリッド分割設定";
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
                Minimum = Constants.MIN_ROWS,
                Maximum = Constants.MAX_ROWS,
                Value = Constants.DEFAULT_ROWS
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
                Minimum = Constants.MIN_COLUMNS,
                Maximum = Constants.MAX_COLUMNS,
                Value = Constants.DEFAULT_COLUMNS
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
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_MARGIN,
                Value = (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES,
                Increment = Constants.INCREMENT_VALUE
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
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_MARGIN,
                Value = (decimal)Constants.DEFAULT_VERTICAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES,
                Increment = Constants.INCREMENT_VALUE
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
    /// 複数図形のグリッド分割設定用ダイアログ
    /// </summary>
    public partial class GridDivisionDialog : Form
    {
        public int Rows { get; private set; }
        public int Columns { get; private set; }
        public float HorizontalMargin { get; private set; }
        public float VerticalMargin { get; private set; }
        public bool DeleteOriginalShapes { get; private set; }

        private List<PowerPoint.Shape> targetShapes;
        private NumericUpDown numRows;
        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalMargin;
        private NumericUpDown numVerticalMargin;
        private CheckBox chkLinkMargins;
        private CheckBox chkDeleteOriginal;
        private Button btnOK;
        private Button btnCancel;
        private Label lblInfo;

        public GridDivisionDialog(List<PowerPoint.Shape> shapes)
        {
            targetShapes = shapes;
            InitializeComponent();
            SetDefaultValues();
            UpdateInfoLabel();
        }

        private void SetDefaultValues()
        {
            Rows = Constants.DEFAULT_ROWS;
            Columns = Constants.DEFAULT_COLUMNS;
            HorizontalMargin = Constants.DEFAULT_HORIZONTAL_MARGIN;
            VerticalMargin = Constants.DEFAULT_VERTICAL_MARGIN;
            DeleteOriginalShapes = true;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "グリッド分割設定（複数図形）";
            this.Size = new Size(400, 350);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 情報表示ラベル
            lblInfo = new Label
            {
                Location = new Point(20, 20),
                Size = new Size(350, 40),
                Text = "",
                ForeColor = Color.DarkBlue,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
            };

            // 行数設定
            var lblRows = new Label
            {
                Text = "行数:",
                Location = new Point(20, 70),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numRows = new NumericUpDown
            {
                Location = new Point(120, 68),
                Size = new Size(80, 20),
                Minimum = Constants.MIN_ROWS,
                Maximum = Constants.MAX_ROWS,
                Value = Constants.DEFAULT_ROWS
            };

            // 列数設定
            var lblColumns = new Label
            {
                Text = "列数:",
                Location = new Point(20, 100),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numColumns = new NumericUpDown
            {
                Location = new Point(120, 98),
                Size = new Size(80, 20),
                Minimum = Constants.MIN_COLUMNS,
                Maximum = Constants.MAX_COLUMNS,
                Value = Constants.DEFAULT_COLUMNS
            };

            // 水平マージン設定
            var lblHorizontalMargin = new Label
            {
                Text = "水平マージン:",
                Location = new Point(20, 130),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numHorizontalMargin = new NumericUpDown
            {
                Location = new Point(120, 128),
                Size = new Size(80, 20),
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_MARGIN,
                Value = (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES,
                Increment = Constants.INCREMENT_VALUE
            };

            var lblHorizontalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 130),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 垂直マージン設定
            var lblVerticalMargin = new Label
            {
                Text = "垂直マージン:",
                Location = new Point(20, 160),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numVerticalMargin = new NumericUpDown
            {
                Location = new Point(120, 158),
                Size = new Size(80, 20),
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_MARGIN,
                Value = (decimal)Constants.DEFAULT_VERTICAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES,
                Increment = Constants.INCREMENT_VALUE
            };

            var lblVerticalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 160),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            //// マージン連動チェックボックス
            //chkLinkMargins = new CheckBox
            //{
            //    Text = "水平・垂直マージンを連動",
            //    Location = new Point(20, 190),
            //    Size = new Size(200, 20),
            //    Checked = true
            //};
            //chkLinkMargins.CheckedChanged += ChkLinkMargins_CheckedChanged;

            // 元図形削除チェックボックス
            chkDeleteOriginal = new CheckBox
            {
                Text = "元の図形を削除する",
                Location = new Point(20, 210),
                Size = new Size(200, 20),
                Checked = true,
                ForeColor = Color.DarkRed,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
            };

            // プレビューラベル
            var lblPreview = new Label
            {
                Text = "プレビュー: 3×3 グリッド",
                Location = new Point(20, 240),
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
                Location = new Point(170, 270),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(260, 270),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
            lblInfo,
            lblRows, numRows,
            lblColumns, numColumns,
            lblHorizontalMargin, numHorizontalMargin, lblHorizontalUnit,
            lblVerticalMargin, numVerticalMargin, lblVerticalUnit,
            chkLinkMargins,
            chkDeleteOriginal,
            lblPreview,
            btnOK, btnCancel
        });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void UpdateInfoLabel()
        {
            if (targetShapes != null && targetShapes.Count > 0)
            {
                var bounds = GetShapeGroupBounds(targetShapes);
                lblInfo.Text = $"選択図形: {targetShapes.Count}個\n" +
                              $"範囲: 幅{bounds.Width:F1}pt × 高さ{bounds.Height:F1}pt";
            }
        }

        private ShapeGroupBounds GetShapeGroupBounds(List<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count == 0)
                return new ShapeGroupBounds();

            float minLeft = shapes.Min(s => s.Left);
            float minTop = shapes.Min(s => s.Top);
            float maxRight = shapes.Max(s => s.Left + s.Width);
            float maxBottom = shapes.Max(s => s.Top + s.Height);

            return new ShapeGroupBounds
            {
                Left = minLeft,
                Top = minTop,
                Right = maxRight,
                Bottom = maxBottom,
                Width = maxRight - minLeft,
                Height = maxBottom - minTop
            };
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
            DeleteOriginalShapes = chkDeleteOriginal.Checked;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numRows?.Dispose();
                numColumns?.Dispose();
                numHorizontalMargin?.Dispose();
                numVerticalMargin?.Dispose();
                chkLinkMargins?.Dispose();
                chkDeleteOriginal?.Dispose();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                lblInfo?.Dispose();
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
            Margin = Constants.DEFAULT_HORIZONTAL_MARGIN;
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
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_MARGIN,
                Value = (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES,
                Increment = Constants.INCREMENT_VALUE
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
            Columns = Constants.DEFAULT_GRID_COLUMNS;
            HorizontalSpacing = Constants.DEFAULT_HORIZONTAL_MARGIN;
            VerticalSpacing = Constants.DEFAULT_VERTICAL_MARGIN;
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
                Minimum = Constants.MIN_COLUMNS,
                Maximum = shapeCount,
                Value = Math.Min(Constants.DEFAULT_GRID_COLUMNS, shapeCount)
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
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_SPACING,
                Value = (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES
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
                Minimum = (decimal)Constants.MIN_MARGIN,
                Maximum = (decimal)Constants.MAX_SPACING,
                Value = (decimal)Constants.DEFAULT_VERTICAL_MARGIN,
                DecimalPlaces = Constants.DECIMAL_PLACES
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
            CenterX = Constants.DEFAULT_CENTER_X;
            CenterY = Constants.DEFAULT_CENTER_Y;
            Radius = Constants.DEFAULT_RADIUS;
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
                Minimum = (decimal)Constants.MIN_CENTER_COORDINATE,
                Maximum = (decimal)Constants.MAX_CENTER_COORDINATE,
                Value = (decimal)Constants.DEFAULT_CENTER_X,
                DecimalPlaces = Constants.DECIMAL_PLACES
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
                Minimum = (decimal)Constants.MIN_CENTER_COORDINATE,
                Maximum = (decimal)Constants.MAX_CENTER_COORDINATE,
                Value = (decimal)Constants.DEFAULT_CENTER_Y,
                DecimalPlaces = Constants.DECIMAL_PLACES
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
                Minimum = (decimal)Constants.MIN_RADIUS,
                Maximum = (decimal)Constants.MAX_RADIUS,
                Value = (decimal)Constants.DEFAULT_RADIUS,
                DecimalPlaces = Constants.DECIMAL_PLACES
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