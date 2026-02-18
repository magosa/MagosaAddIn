using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 複数図形のグリッド分割設定用ダイアログ
    /// </summary>
    public partial class GridDivisionDialog : BaseDialog
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
        private CheckBox chkDeleteOriginal;
        private Label lblInfo;
        private Label lblPreview;

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

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int labelWidth = 90;
            const int numericX = 120;
            const int unitX = 210;
            
            // 情報表示ラベル
            lblInfo = CreateInfoLabel("", new Point(DefaultMargin, currentY), new Size(350, 40));
            currentY += 40 + SmallVerticalSpacing;

            // 行数設定
            var lblRows = CreateLabel("行数:", new Point(DefaultMargin, currentY));
            numRows = CreateNumericUpDown(Constants.MIN_ROWS, Constants.MAX_ROWS,
                Constants.DEFAULT_ROWS, new Point(numericX, currentY - 2));
            currentY += StandardVerticalSpacing;

            // 列数設定
            var lblColumns = CreateLabel("列数:", new Point(DefaultMargin, currentY));
            numColumns = CreateNumericUpDown(Constants.MIN_COLUMNS, Constants.MAX_COLUMNS,
                Constants.DEFAULT_COLUMNS, new Point(numericX, currentY - 2));
            currentY += StandardVerticalSpacing;

            // 水平マージン設定
            var lblHorizontalMargin = CreateLabel("水平マージン:", new Point(DefaultMargin, currentY), labelWidth);
            numHorizontalMargin = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, (decimal)Constants.MAX_MARGIN,
                (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN, new Point(numericX, currentY - 2), 2);
            var lblHorizontalUnit = CreateLabel("pt", new Point(unitX, currentY), 20);
            currentY += StandardVerticalSpacing;

            // 垂直マージン設定
            var lblVerticalMargin = CreateLabel("垂直マージン:", new Point(DefaultMargin, currentY), labelWidth);
            numVerticalMargin = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, (decimal)Constants.MAX_MARGIN,
                (decimal)Constants.DEFAULT_VERTICAL_MARGIN, new Point(numericX, currentY - 2), 2);
            var lblVerticalUnit = CreateLabel("pt", new Point(unitX, currentY), 20);
            currentY += StandardVerticalSpacing + 20;

            // 元図形削除チェックボックス
            chkDeleteOriginal = CreateCheckBox("元の図形を削除する", new Point(DefaultMargin, currentY),
                new Size(200, 20), true);
            chkDeleteOriginal.ForeColor = Color.DarkRed;
            chkDeleteOriginal.Font = BoldFont;
            currentY += StandardVerticalSpacing;

            // プレビューラベル
            lblPreview = new Label
            {
                Text = "プレビュー: 3×3 グリッド",
                Location = new Point(DefaultMargin, currentY),
                Size = new Size(200, 20),
                ForeColor = Color.Gray
            };
            currentY += StandardVerticalSpacing;
            
            // ボタンの上にマージンを追加
            currentY += ButtonTopMargin;

            // ボタン位置計算
            int buttonY = currentY;
            int formHeight = CalculateFormHeight(buttonY);
            
            // フォームの基本設定
            ConfigureForm("グリッド分割設定", 380, formHeight);

            // 値変更時のプレビュー更新
            numRows.ValueChanged += (s, e) => UpdatePreview();
            numColumns.ValueChanged += (s, e) => UpdatePreview();

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblInfo,
                lblRows, numRows,
                lblColumns, numColumns,
                lblHorizontalMargin, numHorizontalMargin, lblHorizontalUnit,
                lblVerticalMargin, numVerticalMargin, lblVerticalUnit,
                chkDeleteOriginal,
                lblPreview
            });

            // ボタンを追加
            AddStandardButtons(buttonY, BtnOK_Click);

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

        private void UpdatePreview()
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
                chkDeleteOriginal?.Dispose();
                lblInfo?.Dispose();
                lblPreview?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
