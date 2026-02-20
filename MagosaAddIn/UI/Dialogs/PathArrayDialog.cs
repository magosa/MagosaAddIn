using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// パス配列設定用ダイアログ
    /// </summary>
    public partial class PathArrayDialog : BaseDialog
    {
        public PathArrayOptions Options { get; private set; }

        private NumericUpDown numCount;
        private RadioButton rbEqualSpacing;
        private RadioButton rbCustomSpacing;
        private NumericUpDown numCustomSpacing;
        private CheckBox chkRotateAlongPath;
        private Label lblCustomSpacing;
        private Label lblCustomUnit;

        public PathArrayDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Options = new PathArrayOptions
            {
                Count = Constants.DEFAULT_ARRAY_COUNT,
                EqualSpacing = true,
                CustomSpacing = Constants.DEFAULT_ARRAY_SPACING,
                RotateAlongPath = false
            };
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("パス配列", 320, 260);

            int yPos = InitialTopMargin;

            // 個数
            var lblCount = CreateLabel("個数:", new Point(DefaultMargin, yPos + 2), 80);
            numCount = CreateNumericUpDown((decimal)Constants.MIN_ARRAY_COUNT, 
                (decimal)Constants.MAX_ARRAY_COUNT, (decimal)Constants.DEFAULT_ARRAY_COUNT,
                new Point(DefaultMargin + 90, yPos), 0);
            var lblCountUnit = CreateLabel("個", new Point(DefaultMargin + 180, yPos + 2), 30);
            yPos += StandardVerticalSpacing + 10;

            // 配置方法
            var lblSpacingMode = CreateLabel("配置方法:", new Point(DefaultMargin, yPos + 2), 80);
            yPos += 25;

            rbEqualSpacing = CreateRadioButton("等間隔配置", 
                new Point(DefaultMargin + 20, yPos), new Size(150, 20), true);
            rbEqualSpacing.CheckedChanged += RbSpacingMode_CheckedChanged;
            yPos += 25;

            rbCustomSpacing = CreateRadioButton("カスタム間隔", 
                new Point(DefaultMargin + 20, yPos), new Size(150, 20), false);
            rbCustomSpacing.CheckedChanged += RbSpacingMode_CheckedChanged;
            yPos += 30;

            lblCustomSpacing = CreateLabel("間隔:", new Point(DefaultMargin + 40, yPos + 2), 50);
            numCustomSpacing = CreateNumericUpDown((decimal)Constants.MIN_ARRAY_SPACING, 
                (decimal)Constants.MAX_ARRAY_SPACING, (decimal)Constants.DEFAULT_ARRAY_SPACING,
                new Point(DefaultMargin + 100, yPos), 1);
            lblCustomUnit = CreateLabel("pt", new Point(DefaultMargin + 190, yPos + 2), 30);
            yPos += StandardVerticalSpacing + 10;

            // パスに沿って回転
            chkRotateAlongPath = CreateCheckBox("パスに沿って回転", 
                new Point(DefaultMargin, yPos), new Size(200, 20), false);
            yPos += StandardVerticalSpacing + 10;

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblCount, numCount, lblCountUnit,
                lblSpacingMode, rbEqualSpacing, rbCustomSpacing,
                lblCustomSpacing, numCustomSpacing, lblCustomUnit,
                chkRotateAlongPath
            });

            // ボタンを追加
            AddStandardButtons(yPos, BtnOK_Click);

            // フォーム高さを計算して設定
            this.ClientSize = new Size(320, CalculateFormHeight(yPos));

            // 初期状態を設定
            UpdateControlStates();

            this.ResumeLayout(false);
        }

        private void RbSpacingMode_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            bool customEnabled = rbCustomSpacing.Checked;
            lblCustomSpacing.Enabled = customEnabled;
            numCustomSpacing.Enabled = customEnabled;
            lblCustomUnit.Enabled = customEnabled;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Options = new PathArrayOptions
            {
                Count = (int)numCount.Value,
                EqualSpacing = rbEqualSpacing.Checked,
                CustomSpacing = (float)numCustomSpacing.Value,
                RotateAlongPath = chkRotateAlongPath.Checked
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numCount?.Dispose();
                rbEqualSpacing?.Dispose();
                rbCustomSpacing?.Dispose();
                numCustomSpacing?.Dispose();
                chkRotateAlongPath?.Dispose();
                lblCustomSpacing?.Dispose();
                lblCustomUnit?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
