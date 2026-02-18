using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 図形分割設定用ダイアログ
    /// </summary>
    public partial class DivisionDialog : BaseDialog
    {
        public int Rows { get; private set; }
        public int Columns { get; private set; }
        public float HorizontalMargin { get; private set; }
        public float VerticalMargin { get; private set; }

        private NumericUpDown numRows;
        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalMargin;
        private NumericUpDown numVerticalMargin;
        private Label lblPreview;

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

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int labelWidth = 90;
            const int numericX = 120;
            const int unitX = 210;

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
            currentY += StandardVerticalSpacing + 30;

            // プレビューラベル
            lblPreview = new Label
            {
                Text = "プレビュー: 2×2 グリッド",
                Location = new Point(DefaultMargin, currentY),
                Size = new Size(200, 20),
                ForeColor = Color.Gray
            };
            currentY += StandardVerticalSpacing;
            
            // ボタンの上にマージンを追加
            currentY += ButtonTopMargin;

            // ボタン位置とフォーム高さの計算
            int buttonY = currentY;
            int formHeight = CalculateFormHeight(buttonY);

            // フォームの基本設定
            ConfigureForm("グリッド分割設定", 350, formHeight);

            // 値変更時のプレビュー更新
            numRows.ValueChanged += (s, e) => UpdatePreview();
            numColumns.ValueChanged += (s, e) => UpdatePreview();

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblRows, numRows,
                lblColumns, numColumns,
                lblHorizontalMargin, numHorizontalMargin, lblHorizontalUnit,
                lblVerticalMargin, numVerticalMargin, lblVerticalUnit,
                lblPreview
            });

            // ボタンを追加
            AddStandardButtons(buttonY, BtnOK_Click);

            this.ResumeLayout(false);
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
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numRows?.Dispose();
                numColumns?.Dispose();
                numHorizontalMargin?.Dispose();
                numVerticalMargin?.Dispose();
                lblPreview?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
