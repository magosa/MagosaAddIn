using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// グリッド配置用ダイアログ
    /// </summary>
    public partial class GridArrangementDialog : BaseDialog
    {
        public int Columns { get; private set; }
        public float HorizontalSpacing { get; private set; }
        public float VerticalSpacing { get; private set; }

        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalSpacing;
        private NumericUpDown numVerticalSpacing;

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

            // フォームの基本設定
            ConfigureForm("グリッド配置設定", 300, 200);

            // 列数設定
            var lblColumns = CreateLabel("列数:", new Point(DefaultMargin, 20));
            numColumns = CreateNumericUpDown(Constants.MIN_COLUMNS, shapeCount,
                Math.Min(Constants.DEFAULT_GRID_COLUMNS, shapeCount), new Point(120, 18));

            // 水平間隔設定
            var lblHorizontalSpacing = CreateLabel("水平間隔:", new Point(DefaultMargin, 50));
            numHorizontalSpacing = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, (decimal)Constants.MAX_SPACING,
                (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN, new Point(120, 48), 2);

            // 垂直間隔設定
            var lblVerticalSpacing = CreateLabel("垂直間隔:", new Point(DefaultMargin, 80));
            numVerticalSpacing = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, (decimal)Constants.MAX_SPACING,
                (decimal)Constants.DEFAULT_VERTICAL_MARGIN, new Point(120, 78), 2);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblColumns, numColumns,
                lblHorizontalSpacing, numHorizontalSpacing,
                lblVerticalSpacing, numVerticalSpacing
            });

            // ボタンを追加
            AddStandardButtons(120, BtnOK_Click);

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Columns = (int)numColumns.Value;
            HorizontalSpacing = (float)numHorizontalSpacing.Value;
            VerticalSpacing = (float)numVerticalSpacing.Value;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numColumns?.Dispose();
                numHorizontalSpacing?.Dispose();
                numVerticalSpacing?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
