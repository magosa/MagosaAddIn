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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
            };

            var lblVerticalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 110),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

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
                lblPreview,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
            };

            var lblVerticalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 160),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

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
        public new float Margin { get; private set; }

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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
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
                DecimalPlaces = 2,
                Increment = 0.01m
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

    /// <summary>
    /// 図形選択条件設定用ダイアログ
    /// </summary>
    public partial class ShapeSelectionDialog : Form
    {
        public SelectionCriteria SelectedCriteria { get; private set; }
        public int MatchingShapeCount { get; private set; }

        private RadioButton rbFillColorOnly;
        private RadioButton rbLineStyleOnly;
        private RadioButton rbFillAndLineStyle;
        private RadioButton rbShapeTypeOnly;
        private Button btnOK;
        private Button btnCancel;
        private Label lblPreview;
        private Label lblBaseShapeInfo;

        private PowerPoint.Shape baseShape;
        private ShapeSelector selector;

        public ShapeSelectionDialog(PowerPoint.Shape baseShape)
        {
            this.baseShape = baseShape;
            this.selector = new ShapeSelector();
            InitializeComponent();
            SetDefaultValues();
            UpdateBaseShapeInfo();
            UpdatePreview();
        }

        private void SetDefaultValues()
        {
            SelectedCriteria = SelectionCriteria.FillColorOnly;
            MatchingShapeCount = 0;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "同一書式図形選択";
            this.Size = new Size(400, 350);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(400, 350);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 基準図形情報表示
            lblBaseShapeInfo = new Label
            {
                Text = "基準図形: ",
                Location = new Point(20, 20),
                Size = new Size(350, 40),
                ForeColor = Color.DarkBlue,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
            };

            // 選択条件グループボックス
            var groupCriteria = new GroupBox
            {
                Text = "選択条件",
                Location = new Point(20, 70),
                Size = new Size(350, 150)
            };

            // 塗りのカラーコードが同じもの
            rbFillColorOnly = new RadioButton
            {
                Text = "塗りのカラーコードが同じもの",
                Location = new Point(20, 25),
                Size = new Size(300, 20),
                Checked = true
            };
            rbFillColorOnly.CheckedChanged += RbCriteria_CheckedChanged;

            // 枠線のスタイルが同じもの
            rbLineStyleOnly = new RadioButton
            {
                Text = "枠線のスタイルが同じもの（色・太さ・破線パターン）",
                Location = new Point(20, 50),
                Size = new Size(300, 20)
            };
            rbLineStyleOnly.CheckedChanged += RbCriteria_CheckedChanged;

            // 塗りと枠線のスタイルが同じもの
            rbFillAndLineStyle = new RadioButton
            {
                Text = "塗りと枠線のスタイルが同じもの",
                Location = new Point(20, 75),
                Size = new Size(300, 20)
            };
            rbFillAndLineStyle.CheckedChanged += RbCriteria_CheckedChanged;

            // シェイプの種類が同じもの
            rbShapeTypeOnly = new RadioButton
            {
                Text = "シェイプの種類が同じもの（四角形、円、三角形など）",
                Location = new Point(20, 100),
                Size = new Size(300, 20)
            };
            rbShapeTypeOnly.CheckedChanged += RbCriteria_CheckedChanged;

            groupCriteria.Controls.AddRange(new Control[] {
                rbFillColorOnly, rbLineStyleOnly, rbFillAndLineStyle, rbShapeTypeOnly
            });

            // プレビューラベル
            lblPreview = new Label
            {
                Text = "一致する図形: 0個",
                Location = new Point(20, 230),
                Size = new Size(350, 20),
                ForeColor = Color.Gray,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Italic)
            };

            // ボタン
            btnOK = new Button
            {
                Text = "選択実行",
                Location = new Point(170, 270),
                Size = new Size(80, 25),
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
                lblBaseShapeInfo,
                groupCriteria,
                lblPreview,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void UpdateBaseShapeInfo()
        {
            if (baseShape != null)
            {
                try
                {
                    var shapeName = ComExceptionHandler.HandleComOperation(
                        () => baseShape.Name,
                        "基準図形名取得",
                        defaultValue: "不明な図形",
                        throwOnError: false);

                    lblBaseShapeInfo.Text = $"基準図形: {shapeName}";
                }
                catch
                {
                    lblBaseShapeInfo.Text = "基準図形: 情報取得エラー";
                }
            }
        }

        private void RbCriteria_CheckedChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            try
            {
                var criteria = GetSelectedCriteria();
                var count = selector.GetMatchingShapeCount(baseShape, criteria);
                MatchingShapeCount = count;

                lblPreview.Text = $"一致する図形: {count}個";
                lblPreview.ForeColor = count > 0 ? Color.DarkGreen : Color.Gray;

                // OKボタンの有効/無効制御
                btnOK.Enabled = count > 0;
            }
            catch (Exception ex)
            {
                lblPreview.Text = "プレビュー取得エラー";
                lblPreview.ForeColor = Color.Red;
                btnOK.Enabled = false;
                ComExceptionHandler.LogError("選択プレビュー更新", ex);
            }
        }

        private SelectionCriteria GetSelectedCriteria()
        {
            if (rbFillColorOnly.Checked)
                return SelectionCriteria.FillColorOnly;
            else if (rbLineStyleOnly.Checked)
                return SelectionCriteria.LineStyleOnly;
            else if (rbFillAndLineStyle.Checked)
                return SelectionCriteria.FillAndLineStyle;
            else if (rbShapeTypeOnly.Checked)
                return SelectionCriteria.ShapeTypeOnly;
            else
                return SelectionCriteria.FillColorOnly;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedCriteria = GetSelectedCriteria();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                rbFillColorOnly?.Dispose();
                rbLineStyleOnly?.Dispose();
                rbFillAndLineStyle?.Dispose();
                rbShapeTypeOnly?.Dispose();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                lblPreview?.Dispose();
                lblBaseShapeInfo?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 動的角度ハンドル設定用ダイアログ
    /// </summary>
    public partial class DynamicAngleHandleDialog : Form
    {
        public float[] HandleValues { get; private set; }
        public bool DialogResult_OK { get; private set; }

        private List<NumericUpDown> handleControls;
        private List<Label> interpretationLabels;
        private Button btnOK;
        private Button btnCancel;
        private Button btnGetCurrentValues;
        private Label lblShapeInfo;
        private GroupBox groupAngleHandles;

        private List<PowerPoint.Shape> targetShapes;
        private ShapeHandleAdjuster adjuster;
        private ShapeHandleAnalysis analysis;

        public DynamicAngleHandleDialog(List<PowerPoint.Shape> shapes, ShapeHandleAnalysis analysis)
        {
            targetShapes = shapes;
            this.analysis = analysis;
            adjuster = new ShapeHandleAdjuster();
            handleControls = new List<NumericUpDown>();
            interpretationLabels = new List<Label>();

            InitializeComponent();
            CreateDynamicAngleControls();
            UpdateShapeInfo();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "角度ハンドル設定";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 図形情報表示
            lblShapeInfo = new Label
            {
                Location = new Point(20, 20),
                Size = new Size(450, 80),
                Text = "図形情報を取得中...",
                ForeColor = Color.DarkBlue,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
            };

            // 現在値取得ボタン
            btnGetCurrentValues = new Button
            {
                Text = "現在の値を取得",
                Location = new Point(20, 110),
                Size = new Size(120, 25)
            };
            btnGetCurrentValues.Click += BtnGetCurrentValues_Click;

            // 角度ハンドルグループ
            groupAngleHandles = new GroupBox
            {
                Text = "角度ハンドル値（度数）",
                Location = new Point(20, 150),
                Size = new Size(450, 200) // 動的に調整
            };

            // 基本コントロールを追加
            this.Controls.AddRange(new Control[] {
                lblShapeInfo,
                btnGetCurrentValues,
                groupAngleHandles
            });

            this.ResumeLayout(false);
        }

        private void CreateDynamicAngleControls()
        {
            if (analysis == null || analysis.RecommendedAngleHandleCount == 0)
            {
                // 角度ハンドルがない場合
                var lblNoHandles = new Label
                {
                    Text = "選択された図形には角度ハンドルがありません。\n" +
                           "角度ハンドル対応図形: 円弧、弦、扇形、ブロック円弧、ドーナツ、三日月など",
                    Location = new Point(20, 30),
                    Size = new Size(400, 40),
                    ForeColor = Color.Gray
                };
                groupAngleHandles.Controls.Add(lblNoHandles);

                // フォームサイズを調整
                this.Size = new Size(500, 320);

                // ボタンを追加
                AddButtons(270);
                return;
            }

            // 角度ハンドル図形の代表例を取得
            var representativeShape = analysis.ShapeInfos
                .FirstOrDefault(info => info.IsAngleHandleShape && info.AdjustmentCount > 0);

            if (representativeShape == null)
            {
                CreateDynamicAngleControls(); // 再帰的に呼び出し（角度ハンドルなしとして処理）
                return;
            }

            // 動的に角度ハンドルコントロールを作成
            int handleCount = Math.Min(representativeShape.AdjustmentCount, Constants.MAX_SUPPORTED_HANDLES);
            HandleValues = new float[handleCount];

            for (int i = 0; i < handleCount; i++)
            {
                // ハンドルの意味を表示
                string handleMeaning = i < representativeShape.AngleInterpretation.Count
                    ? representativeShape.AngleInterpretation[i]
                    : $"ハンドル{i + 1}";

                var lblHandle = new Label
                {
                    Text = $"{handleMeaning}:",
                    Location = new Point(20, 30 + i * 50),
                    Size = new Size(100, 20),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
                };

                var numHandle = new NumericUpDown
                {
                    Location = new Point(130, 28 + i * 50),
                    Size = new Size(100, 20),
                    Minimum = (decimal)Constants.MIN_ANGLE_DEGREE,
                    Maximum = (decimal)Constants.MAX_ANGLE_DEGREE,
                    Value = (decimal)GetInitialAngleValue(i),
                    DecimalPlaces = 2,
                    Increment = 0.01m,
                    Tag = i // インデックスを保存
                };

                var lblUnit = new Label
                {
                    Text = "°",
                    Location = new Point(240, 30 + i * 50),
                    Size = new Size(20, 20),
                    ForeColor = Color.Gray
                };

                // 角度の説明ラベル
                var lblInterpretation = new Label
                {
                    Text = GetAngleDescription(representativeShape.ShapeType, i),
                    Location = new Point(270, 30 + i * 50),
                    Size = new Size(160, 20),
                    ForeColor = Color.Gray,
                    Font = new Font(SystemFonts.DefaultFont.FontFamily, 8)
                };

                handleControls.Add(numHandle);
                interpretationLabels.Add(lblInterpretation);
                groupAngleHandles.Controls.AddRange(new Control[] { lblHandle, numHandle, lblUnit, lblInterpretation });
            }

            // グループボックスのサイズを調整
            int groupHeight = Math.Max(100, 60 + handleCount * 50);
            groupAngleHandles.Size = new Size(450, groupHeight);

            // フォームサイズを調整
            int formHeight = 220 + groupHeight + 80;
            this.Size = new Size(500, formHeight);

            // ボタンを追加
            AddButtons(formHeight - 80);
        }

        /// <summary>
        /// 初期角度値を取得
        /// </summary>
        /// <param name="handleIndex">ハンドルインデックス</param>
        /// <returns>初期角度値</returns>
        private float GetInitialAngleValue(int handleIndex)
        {
            try
            {
                if (targetShapes != null && targetShapes.Count > 0)
                {
                    var firstAngleShape = targetShapes
                        .FirstOrDefault(s => adjuster.GetHandleInfoFast(s).IsAngleHandleShape);

                    if (firstAngleShape != null && handleIndex < firstAngleShape.Adjustments.Count)
                    {
                        float currentValue = firstAngleShape.Adjustments[handleIndex + 1]; // PowerPointは1ベース
                        float degreeValue = adjuster.ConvertNormalizedToDegreeByShapeType(currentValue, firstAngleShape.AutoShapeType, handleIndex);

                        ComExceptionHandler.LogDebug($"初期値取得: ハンドル{handleIndex + 1} = {currentValue} → {degreeValue}°");

                        return Math.Max(Constants.MIN_ANGLE_DEGREE, Math.Min(Constants.MAX_ANGLE_DEGREE, degreeValue));
                    }
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError($"初期値取得エラー: ハンドル{handleIndex + 1}", ex);
            }

            return Constants.DEFAULT_ANGLE_DEGREE;
        }

        private string GetAngleDescription(string shapeType, int handleIndex)
        {
            switch (shapeType)
            {
                case "msoShapeArc":
                    return handleIndex == 0 ? "円弧の開始位置" : "円弧の終了位置";
                case "msoShapeChord":
                    return handleIndex == 0 ? "弦の開始角度" : "弦の終了角度";
                case "msoShapePie":
                    return handleIndex == 0 ? "扇形の開始角度" : "扇形の終了角度";
                case "msoShapeBlockArc":
                    if (handleIndex == 0) return "ブロック円弧の開始角度";
                    if (handleIndex == 1) return "ブロック円弧の終了角度";
                    return "内径の比率";
                case "msoShapeDonut":
                    return "ドーナツの内径比率";
                case "msoShapeMoon":
                    return "三日月の角度";
                default:
                    return $"角度調整値{handleIndex + 1}";
            }
        }

        private void AddButtons(int yPosition)
        {
            btnOK = new Button
            {
                Text = "適用",
                Location = new Point(270, yPosition),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(360, yPosition),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] { btnOK, btnCancel });
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            // 角度ハンドルがない場合はOKボタンを無効化
            if (analysis?.RecommendedAngleHandleCount == 0)
            {
                btnOK.Enabled = false;
            }
        }

        private void UpdateShapeInfo()
        {
            if (analysis != null)
            {
                var angleShapeTypes = analysis.ShapeInfos
                    .Where(info => info.IsAngleHandleShape)
                    .Select(info => info.GetDisplayShapeType())
                    .Distinct()
                    .ToList();

                lblShapeInfo.Text = $"選択図形: {analysis.TotalShapes}個\n" +
                                   $"角度ハンドル対応図形: {analysis.ShapesWithAngleHandles}個\n" +
                                   $"図形タイプ: {string.Join(", ", angleShapeTypes)}";

                if (analysis.RecommendedAngleHandleCount == 0)
                {
                    lblShapeInfo.Text += "\n※角度ハンドル対応図形を選択してください";
                    lblShapeInfo.ForeColor = Color.DarkRed;
                }
            }
        }

        private void BtnGetCurrentValues_Click(object sender, EventArgs e)
        {
            if (targetShapes != null && targetShapes.Count > 0 && handleControls.Count > 0)
            {
                try
                {
                    // 最初の角度ハンドル図形の現在値を取得
                    var firstAngleShape = targetShapes
                        .FirstOrDefault(s => adjuster.GetHandleInfoFast(s).IsAngleHandleShape);

                    if (firstAngleShape != null)
                    {
                        ComExceptionHandler.LogDebug("=== 角度ハンドル現在値取得開始 ===");

                        for (int i = 0; i < Math.Min(handleControls.Count, firstAngleShape.Adjustments.Count); i++)
                        {
                            float currentValue = firstAngleShape.Adjustments[i + 1]; // PowerPointは1ベース
                            ComExceptionHandler.LogDebug($"現在のハンドル{i + 1}値: {currentValue}");

                            // 図形タイプに応じて度数に変換
                            float degreeValue = adjuster.ConvertNormalizedToDegreeByShapeType(currentValue, firstAngleShape.AutoShapeType, i);
                            ComExceptionHandler.LogDebug($"度数変換: {currentValue} → {degreeValue}°");

                            // 範囲内にクランプ
                            decimal clampedValue = (decimal)Math.Max(Constants.MIN_ANGLE_DEGREE,
                                Math.Min(Constants.MAX_ANGLE_DEGREE, degreeValue));

                            handleControls[i].Value = clampedValue;
                            ComExceptionHandler.LogDebug($"ダイアログ設定: {clampedValue}°");
                        }

                        var shapeInfo = adjuster.GetHandleInfoFast(firstAngleShape);
                        MessageBox.Show($"図形「{firstAngleShape.Name}」({shapeInfo.GetDisplayShapeType()})の現在値を取得しました。",
                            "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("角度ハンドルを持つ図形が見つかりません。",
                            "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.ShowOperationError("現在値取得", ex);
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (HandleValues != null)
            {
                for (int i = 0; i < HandleValues.Length && i < handleControls.Count; i++)
                {
                    HandleValues[i] = (float)handleControls[i].Value;
                }
            }
            DialogResult_OK = true;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                handleControls?.ForEach(control => control?.Dispose());
                handleControls?.Clear();
                interpretationLabels?.ForEach(label => label?.Dispose());
                interpretationLabels?.Clear();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                btnGetCurrentValues?.Dispose();
                lblShapeInfo?.Dispose();
                groupAngleHandles?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 動的調整ハンドル設定用ダイアログ（完全版）
    /// </summary>
    public partial class DynamicHandleDialog : Form
    {
        public float[] HandleValues { get; private set; }
        public bool DialogResult_OK { get; private set; }

        private List<NumericUpDown> handleControls;
        private Button btnOK;
        private Button btnCancel;
        private Button btnGetCurrentValues;
        private Label lblShapeInfo;
        private GroupBox groupHandles;

        private List<PowerPoint.Shape> targetShapes;
        private ShapeHandleAdjuster adjuster;
        private ShapeHandleAnalysis analysis;

        public DynamicHandleDialog(List<PowerPoint.Shape> shapes, ShapeHandleAnalysis analysis)
        {
            targetShapes = shapes;
            this.analysis = analysis;
            adjuster = new ShapeHandleAdjuster();
            handleControls = new List<NumericUpDown>();

            InitializeComponent();
            CreateDynamicControls();
            UpdateShapeInfo();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "調整ハンドル設定";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 図形情報表示
            lblShapeInfo = new Label
            {
                Location = new Point(20, 20),
                Size = new Size(400, 60),
                Text = "図形情報を取得中...",
                ForeColor = Color.DarkBlue,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
            };

            // 現在値取得ボタン
            btnGetCurrentValues = new Button
            {
                Text = "現在の値を取得",
                Location = new Point(20, 90),
                Size = new Size(120, 25)
            };
            btnGetCurrentValues.Click += BtnGetCurrentValues_Click;

            // 調整ハンドルグループ
            groupHandles = new GroupBox
            {
                Text = "調整ハンドル値（mm単位）",
                Location = new Point(20, 130),
                Size = new Size(400, 200) // 動的に調整
            };

            // 基本コントロールを追加
            this.Controls.AddRange(new Control[] {
                lblShapeInfo,
                btnGetCurrentValues,
                groupHandles
            });

            this.ResumeLayout(false);
        }

        private void CreateDynamicControls()
        {
            if (analysis == null || analysis.RecommendedHandleCount == 0)
            {
                // 調整ハンドルがない場合
                var lblNoHandles = new Label
                {
                    Text = "選択された図形には調整ハンドルがありません。\n\n" +
                           "調整ハンドル対応図形の例:\n" +
                           "・角丸四角形（角丸の調整）\n" +
                           "・吹き出し（尻尾の位置調整）\n" +
                           "・矢印（矢じりの調整）\n" +
                           "・星形（内側の頂点調整）など",
                    Location = new Point(20, 30),
                    Size = new Size(350, 100),
                    ForeColor = Color.Gray
                };
                groupHandles.Controls.Add(lblNoHandles);

                // フォームサイズを調整
                this.Size = new Size(450, 350);

                // ボタンを追加
                AddButtons(300);
                return;
            }

            // 動的にハンドルコントロールを作成
            int handleCount = Math.Min(analysis.RecommendedHandleCount, Constants.MAX_SUPPORTED_HANDLES);
            HandleValues = new float[handleCount];

            // 調整ハンドル図形かどうかを判定（変数名を変更）
            bool useAdjustmentHandleUnit = analysis.ShapeInfos.Any(info => info.IsAdjustmentHandleShape && !info.IsAngleHandleShape);

            for (int i = 0; i < handleCount; i++)
            {
                var lblHandle = new Label
                {
                    Text = $"ハンドル {i + 1}:",
                    Location = new Point(20, 30 + i * 35),
                    Size = new Size(80, 20),
                    TextAlign = ContentAlignment.MiddleLeft
                };

                var numHandle = new NumericUpDown
                {
                    Location = new Point(110, 28 + i * 35),
                    Size = new Size(100, 20),
                    DecimalPlaces = 2,
                    Increment = 0.01m,
                    Tag = i // インデックスを保存
                };

                // 単位に応じて設定を変更
                if (useAdjustmentHandleUnit)
                {
                    // mm単位での設定
                    numHandle.Minimum = (decimal)Constants.MIN_HANDLE_MM;
                    numHandle.Maximum = (decimal)Constants.MAX_HANDLE_MM;
                    numHandle.Value = (decimal)GetInitialHandleValue(i);
                }
                else
                {
                    // 従来の正規化値（0.0-1.0）
                    numHandle.Minimum = (decimal)Constants.MIN_HANDLE_VALUE;
                    numHandle.Maximum = (decimal)Constants.MAX_HANDLE_VALUE;
                    numHandle.Value = (decimal)GetInitialHandleValue(i);
                    numHandle.DecimalPlaces = 3;
                }

                var lblUnit = new Label
                {
                    Text = useAdjustmentHandleUnit ? "mm" : "",
                    Location = new Point(220, 30 + i * 35),
                    Size = new Size(30, 20),
                    ForeColor = Color.Gray
                };

                var lblDescription = new Label
                {
                    Text = GetHandleDescription(i),
                    Location = new Point(260, 30 + i * 35),
                    Size = new Size(120, 20),
                    ForeColor = Color.Gray,
                    Font = new Font(SystemFonts.DefaultFont.FontFamily, 8)
                };

                handleControls.Add(numHandle);
                groupHandles.Controls.AddRange(new Control[] { lblHandle, numHandle, lblUnit, lblDescription });
            }

            // グループボックスのサイズを調整
            int groupHeight = Math.Max(100, 60 + handleCount * 35);
            groupHandles.Size = new Size(400, groupHeight);

            // フォームサイズを調整
            int formHeight = 200 + groupHeight + 80;
            this.Size = new Size(450, formHeight);

            // ボタンを追加
            AddButtons(formHeight - 80);
        }

        /// <summary>
        /// 初期調整ハンドル値を取得
        /// </summary>
        /// <param name="handleIndex">ハンドルインデックス</param>
        /// <returns>初期調整ハンドル値</returns>
        private float GetInitialHandleValue(int handleIndex)
        {
            try
            {
                if (targetShapes != null && targetShapes.Count > 0)
                {
                    var firstShapeWithHandles = targetShapes
                        .FirstOrDefault(s => adjuster.GetHandleInfoFast(s).AdjustmentCount > 0);

                    if (firstShapeWithHandles != null && handleIndex < firstShapeWithHandles.Adjustments.Count)
                    {
                        float currentValue = firstShapeWithHandles.Adjustments[handleIndex + 1]; // PowerPointは1ベース

                        // 調整ハンドル図形かどうかを判定（変数名を変更）
                        var shapeInfo = adjuster.GetHandleInfoFast(firstShapeWithHandles);
                        bool shouldUseMillimeterUnit = shapeInfo.IsAdjustmentHandleShape && !shapeInfo.IsAngleHandleShape;

                        if (shouldUseMillimeterUnit)
                        {
                            // mm単位の場合
                            float mmValue = adjuster.ConvertNormalizedToMm(currentValue, firstShapeWithHandles, handleIndex);
                            ComExceptionHandler.LogDebug($"初期値取得(mm): ハンドル{handleIndex + 1} = {currentValue} → {mmValue}mm");
                            return Math.Max(Constants.MIN_HANDLE_MM, Math.Min(Constants.MAX_HANDLE_MM, mmValue));
                        }
                        else
                        {
                            // 正規化値の場合
                            ComExceptionHandler.LogDebug($"初期値取得(正規化): ハンドル{handleIndex + 1} = {currentValue}");
                            return Math.Max(Constants.MIN_HANDLE_VALUE, Math.Min(Constants.MAX_HANDLE_VALUE, currentValue));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError($"初期値取得エラー: ハンドル{handleIndex + 1}", ex);
            }

            // デフォルト値を返す（変数名を変更）
            bool shouldReturnMmDefault = analysis?.ShapeInfos.Any(info => info.IsAdjustmentHandleShape && !info.IsAngleHandleShape) ?? false;
            return shouldReturnMmDefault ? Constants.DEFAULT_HANDLE_MM : Constants.DEFAULT_HANDLE_VALUE;
        }

        private string GetHandleDescription(int handleIndex)
        {
            // 代表的な図形の調整ハンドルの説明
            return $"調整値 {handleIndex + 1}";
        }

        private void AddButtons(int yPosition)
        {
            btnOK = new Button
            {
                Text = "適用",
                Location = new Point(220, yPosition),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(310, yPosition),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] { btnOK, btnCancel });
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            // 調整ハンドルがない場合はOKボタンを無効化
            if (analysis?.RecommendedHandleCount == 0)
            {
                btnOK.Enabled = false;
            }
        }

        private void UpdateShapeInfo()
        {
            if (analysis != null)
            {
                lblShapeInfo.Text = $"選択図形: {analysis.TotalShapes}個\n" +
                                   $"調整ハンドル有り: {analysis.ShapesWithAdjustmentHandles}個\n" +
                                   $"推奨ハンドル数: {analysis.RecommendedHandleCount}個";

                if (analysis.RecommendedHandleCount == 0)
                {
                    lblShapeInfo.Text += "\n※調整ハンドルを持つ図形を選択してください";
                    lblShapeInfo.ForeColor = Color.DarkRed;
                }
            }
        }

        private void BtnGetCurrentValues_Click(object sender, EventArgs e)
        {
            if (targetShapes != null && targetShapes.Count > 0 && handleControls.Count > 0)
            {
                try
                {
                    // 最初の図形の現在値を取得
                    var firstShapeWithHandles = targetShapes
                        .FirstOrDefault(s => adjuster.GetHandleInfoFast(s).AdjustmentCount > 0);

                    if (firstShapeWithHandles != null)
                    {
                        ComExceptionHandler.LogDebug("=== 調整ハンドル現在値取得開始 ===");

                        // 調整ハンドル図形かどうかを判定（変数名を変更）
                        var shapeInfo = adjuster.GetHandleInfoFast(firstShapeWithHandles);
                        bool shouldDisplayInMillimeter = shapeInfo.IsAdjustmentHandleShape && !shapeInfo.IsAngleHandleShape;

                        ComExceptionHandler.LogDebug($"図形タイプ: {shapeInfo.ShapeType}, 調整ハンドル: {shouldDisplayInMillimeter}");

                        for (int i = 0; i < Math.Min(handleControls.Count, firstShapeWithHandles.Adjustments.Count); i++)
                        {
                            float currentValue = firstShapeWithHandles.Adjustments[i + 1]; // PowerPointは1ベース
                            ComExceptionHandler.LogDebug($"現在のハンドル{i + 1}値: {currentValue}");

                            if (shouldDisplayInMillimeter)
                            {
                                // mm単位の調整ハンドルの場合
                                float mmValue = adjuster.ConvertNormalizedToMm(currentValue, firstShapeWithHandles, i);
                                ComExceptionHandler.LogDebug($"mm変換: {currentValue} → {mmValue}mm");

                                // mm値を範囲内にクランプ
                                decimal clampedValue = (decimal)Math.Max(Constants.MIN_HANDLE_MM,
                                    Math.Min(Constants.MAX_HANDLE_MM, mmValue));

                                handleControls[i].Value = clampedValue;
                                ComExceptionHandler.LogDebug($"ダイアログ設定: {clampedValue}mm");
                            }
                            else
                            {
                                // 正規化値（0.0-1.0）の調整ハンドルの場合
                                ComExceptionHandler.LogDebug($"正規化値使用: {currentValue}");

                                // 正規化値を範囲内にクランプ
                                decimal clampedValue = (decimal)Math.Max(Constants.MIN_HANDLE_VALUE,
                                    Math.Min(Constants.MAX_HANDLE_VALUE, currentValue));

                                handleControls[i].Value = clampedValue;
                                ComExceptionHandler.LogDebug($"ダイアログ設定: {clampedValue}");
                            }
                        }

                        string unitText = shouldDisplayInMillimeter ? "mm単位" : "正規化値";
                        MessageBox.Show($"図形「{firstShapeWithHandles.Name}」の現在値を取得しました。（{unitText}）",
                            "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("調整ハンドルを持つ図形が見つかりません。",
                            "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.ShowOperationError("現在値取得", ex);
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (HandleValues != null)
            {
                for (int i = 0; i < HandleValues.Length && i < handleControls.Count; i++)
                {
                    HandleValues[i] = (float)handleControls[i].Value;
                }
            }
            DialogResult_OK = true;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                handleControls?.ForEach(control => control?.Dispose());
                handleControls?.Clear();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                btnGetCurrentValues?.Dispose();
                lblShapeInfo?.Dispose();
                groupHandles?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}