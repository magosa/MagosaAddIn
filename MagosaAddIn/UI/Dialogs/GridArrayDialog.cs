using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// グリッド配列設定用ダイアログ
    /// </summary>
    public partial class GridArrayDialog : BaseDialog
    {
        public GridArrayOptions Options { get; private set; }

        private NumericUpDown numRows;
        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalSpacing;
        private NumericUpDown numVerticalSpacing;
        private NumericUpDown numAngle;

        public GridArrayDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Options = new GridArrayOptions
            {
                Rows = Constants.DEFAULT_ROWS,
                Columns = Constants.DEFAULT_COLUMNS,
                HorizontalSpacing = Constants.DEFAULT_HORIZONTAL_MARGIN,
                VerticalSpacing = Constants.DEFAULT_VERTICAL_MARGIN
            };
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("グリッド配列", 320, 240);

            int yPos = InitialTopMargin;

            // 行数
            var lblRows = CreateLabel("行数:", new Point(DefaultMargin, yPos + 2), 90);
            numRows = CreateNumericUpDown((decimal)Constants.MIN_ROWS, 
                (decimal)Constants.MAX_ROWS, (decimal)Constants.DEFAULT_ROWS,
                new Point(DefaultMargin + 100, yPos), 0);
            var lblRowsUnit = CreateLabel("行", new Point(DefaultMargin + 190, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 列数
            var lblColumns = CreateLabel("列数:", new Point(DefaultMargin, yPos + 2), 90);
            numColumns = CreateNumericUpDown((decimal)Constants.MIN_COLUMNS, 
                (decimal)Constants.MAX_COLUMNS, (decimal)Constants.DEFAULT_COLUMNS,
                new Point(DefaultMargin + 100, yPos), 0);
            var lblColumnsUnit = CreateLabel("列", new Point(DefaultMargin + 190, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 水平間隔
            var lblHorizontalSpacing = CreateLabel("水平間隔:", new Point(DefaultMargin, yPos + 2), 90);
            numHorizontalSpacing = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, 
                (decimal)Constants.MAX_SPACING, (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN,
                new Point(DefaultMargin + 100, yPos), 1);
            var lblHorizontalUnit = CreateLabel("pt", new Point(DefaultMargin + 190, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 垂直間隔
            var lblVerticalSpacing = CreateLabel("垂直間隔:", new Point(DefaultMargin, yPos + 2), 90);
            numVerticalSpacing = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, 
                (decimal)Constants.MAX_SPACING, (decimal)Constants.DEFAULT_VERTICAL_MARGIN,
                new Point(DefaultMargin + 100, yPos), 1);
            var lblVerticalUnit = CreateLabel("pt", new Point(DefaultMargin + 190, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 角度
            var lblAngle = CreateLabel("角度:", new Point(DefaultMargin, yPos + 2), 90);
            numAngle = CreateNumericUpDown(-360, 360, 0,
                new Point(DefaultMargin + 100, yPos), 1);
            var lblAngleUnit = CreateLabel("°", new Point(DefaultMargin + 190, yPos + 2), 30);
            var lblAngleHint = CreateLabel("(0°=右, 90°=下)", new Point(DefaultMargin, yPos + 25), 200);
            lblAngleHint.Font = new Font(lblAngleHint.Font.FontFamily, 7.5f);
            yPos += StandardVerticalSpacing + 10;

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblRows, numRows, lblRowsUnit,
                lblColumns, numColumns, lblColumnsUnit,
                lblHorizontalSpacing, numHorizontalSpacing, lblHorizontalUnit,
                lblVerticalSpacing, numVerticalSpacing, lblVerticalUnit,
                lblAngle, numAngle, lblAngleUnit, lblAngleHint
            });

            // ボタンを追加
            AddStandardButtons(yPos, BtnOK_Click);

            // フォーム高さを計算して設定
            this.ClientSize = new Size(320, CalculateFormHeight(yPos));

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Options = new GridArrayOptions
            {
                Rows = (int)numRows.Value,
                Columns = (int)numColumns.Value,
                HorizontalSpacing = (float)numHorizontalSpacing.Value,
                VerticalSpacing = (float)numVerticalSpacing.Value,
                Angle = (float)numAngle.Value
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numRows?.Dispose();
                numColumns?.Dispose();
                numHorizontalSpacing?.Dispose();
                numVerticalSpacing?.Dispose();
                numAngle?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
