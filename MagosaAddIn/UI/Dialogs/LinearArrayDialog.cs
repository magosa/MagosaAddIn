using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 線形配列設定用ダイアログ
    /// </summary>
    public partial class LinearArrayDialog : BaseDialog
    {
        public LinearArrayOptions Options { get; private set; }

        private NumericUpDown numAngle;
        private NumericUpDown numCount;
        private NumericUpDown numSpacing;

        public LinearArrayDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Options = new LinearArrayOptions
            {
                Angle = Constants.DEFAULT_LINEAR_ANGLE,
                Count = Constants.DEFAULT_ARRAY_COUNT,
                Spacing = Constants.DEFAULT_ARRAY_SPACING
            };
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("線形配列", 300, 230);

            int yPos = InitialTopMargin;

            // 配列方向（角度）
            var lblAngle = CreateLabel("配列方向:", new Point(DefaultMargin, yPos + 2), 80);
            numAngle = CreateNumericUpDown((decimal)Constants.MIN_ROTATION_ANGLE, 
                (decimal)Constants.MAX_ROTATION_ANGLE, (decimal)Constants.DEFAULT_LINEAR_ANGLE,
                new Point(DefaultMargin + 90, yPos), 1);
            var lblAngleUnit = CreateLabel("度", new Point(DefaultMargin + 180, yPos + 2), 30);
            var lblAngleNote = CreateLabel("(0°=右, 90°=下)", new Point(DefaultMargin + 90, yPos + 25), 120);
            lblAngleNote.Font = SmallFont;
            lblAngleNote.ForeColor = Color.Gray;
            yPos += StandardVerticalSpacing + 20;

            // 個数
            var lblCount = CreateLabel("個数:", new Point(DefaultMargin, yPos + 2), 80);
            numCount = CreateNumericUpDown((decimal)Constants.MIN_ARRAY_COUNT, 
                (decimal)Constants.MAX_ARRAY_COUNT, (decimal)Constants.DEFAULT_ARRAY_COUNT,
                new Point(DefaultMargin + 90, yPos), 0);
            var lblCountUnit = CreateLabel("個", new Point(DefaultMargin + 180, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 間隔
            var lblSpacing = CreateLabel("間隔:", new Point(DefaultMargin, yPos + 2), 80);
            numSpacing = CreateNumericUpDown((decimal)Constants.MIN_ARRAY_SPACING, 
                (decimal)Constants.MAX_ARRAY_SPACING, (decimal)Constants.DEFAULT_ARRAY_SPACING,
                new Point(DefaultMargin + 90, yPos), 1);
            var lblSpacingUnit = CreateLabel("pt", new Point(DefaultMargin + 180, yPos + 2), 30);
            yPos += StandardVerticalSpacing + 10;

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblAngle, numAngle, lblAngleUnit, lblAngleNote,
                lblCount, numCount, lblCountUnit,
                lblSpacing, numSpacing, lblSpacingUnit
            });

            // ボタンを追加
            AddStandardButtons(yPos, BtnOK_Click);

            // フォーム高さを計算して設定
            this.ClientSize = new Size(300, CalculateFormHeight(yPos));

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Options = new LinearArrayOptions
            {
                Angle = (float)numAngle.Value,
                Count = (int)numCount.Value,
                Spacing = (float)numSpacing.Value
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numAngle?.Dispose();
                numCount?.Dispose();
                numSpacing?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
