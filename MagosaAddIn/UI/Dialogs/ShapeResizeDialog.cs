using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 図形サイズ調整用ダイアログ
    /// </summary>
    public partial class ShapeResizeDialog : BaseDialog
    {
        public float Percentage { get; private set; }
        public float Width { get; private set; }
        public float Height { get; private set; }
        public SizeUnit Unit { get; private set; }
        public bool KeepRatio { get; private set; }
        public bool UsePercentage { get; private set; }

        private RadioButton rbPercentage;
        private RadioButton rbFixedSize;
        private NumericUpDown numPercentage;
        private NumericUpDown numWidth;
        private NumericUpDown numHeight;
        private ComboBox cboUnit;
        private CheckBox chkKeepRatio;
        private Label lblPercentageUnit;
        private Label lblWidth;
        private Label lblHeight;
        private Label lblUnit;

        public ShapeResizeDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Percentage = Constants.DEFAULT_PERCENTAGE;
            Width = Constants.DEFAULT_FIXED_WIDTH;
            Height = Constants.DEFAULT_FIXED_HEIGHT;
            Unit = SizeUnit.Point;
            KeepRatio = false;
            UsePercentage = true;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("図形サイズ調整", 350, 280);

            int yPos = InitialTopMargin;

            // パーセント拡大縮小
            rbPercentage = CreateRadioButton("パーセント拡大縮小", 
                new Point(DefaultMargin, yPos), new Size(200, 20), true);
            rbPercentage.CheckedChanged += RbResizeMode_CheckedChanged;
            yPos += StandardVerticalSpacing;

            numPercentage = CreateNumericUpDown((decimal)Constants.MIN_PERCENTAGE, 
                (decimal)Constants.MAX_PERCENTAGE, (decimal)Constants.DEFAULT_PERCENTAGE,
                new Point(DefaultMargin + 20, yPos), 1);
            lblPercentageUnit = CreateLabel("%", new Point(DefaultMargin + 110, yPos + 2), 20);
            yPos += StandardVerticalSpacing + 10;

            // 固定サイズ設定
            rbFixedSize = CreateRadioButton("固定サイズ設定", 
                new Point(DefaultMargin, yPos), new Size(200, 20), false);
            rbFixedSize.CheckedChanged += RbResizeMode_CheckedChanged;
            yPos += StandardVerticalSpacing;

            lblWidth = CreateLabel("幅:", new Point(DefaultMargin + 20, yPos + 2), 40);
            numWidth = CreateNumericUpDown(0.1m, 10000m, (decimal)Constants.DEFAULT_FIXED_WIDTH,
                new Point(DefaultMargin + 65, yPos), 1);
            yPos += StandardVerticalSpacing;

            lblHeight = CreateLabel("高さ:", new Point(DefaultMargin + 20, yPos + 2), 40);
            numHeight = CreateNumericUpDown(0.1m, 10000m, (decimal)Constants.DEFAULT_FIXED_HEIGHT,
                new Point(DefaultMargin + 65, yPos), 1);
            yPos += StandardVerticalSpacing;

            lblUnit = CreateLabel("単位:", new Point(DefaultMargin + 20, yPos + 2), 40);
            cboUnit = CreateComboBox(new Point(DefaultMargin + 65, yPos), new Size(100, 20));
            cboUnit.Items.AddRange(new object[] { "ポイント (pt)", "ミリメートル (mm)", "センチメートル (cm)" });
            cboUnit.SelectedIndex = 0;
            yPos += StandardVerticalSpacing;

            chkKeepRatio = CreateCheckBox("アスペクト比を保持", 
                new Point(DefaultMargin + 20, yPos), new Size(150, 20), false);
            yPos += StandardVerticalSpacing + 10;

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                rbPercentage, numPercentage, lblPercentageUnit,
                rbFixedSize, lblWidth, numWidth, lblHeight, numHeight, lblUnit, cboUnit, chkKeepRatio
            });

            // ボタンを追加
            int buttonY = yPos;
            AddStandardButtons(buttonY, BtnOK_Click);

            // フォーム高さを計算して設定
            this.ClientSize = new Size(350, CalculateFormHeight(buttonY));

            // 初期状態を設定
            UpdateControlStates();

            this.ResumeLayout(false);
        }

        private void RbResizeMode_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            bool isPercentage = rbPercentage.Checked;

            numPercentage.Enabled = isPercentage;
            lblPercentageUnit.Enabled = isPercentage;

            numWidth.Enabled = !isPercentage;
            numHeight.Enabled = !isPercentage;
            lblWidth.Enabled = !isPercentage;
            lblHeight.Enabled = !isPercentage;
            lblUnit.Enabled = !isPercentage;
            cboUnit.Enabled = !isPercentage;
            chkKeepRatio.Enabled = !isPercentage;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            UsePercentage = rbPercentage.Checked;

            if (UsePercentage)
            {
                Percentage = (float)numPercentage.Value;
            }
            else
            {
                Width = (float)numWidth.Value;
                Height = (float)numHeight.Value;
                Unit = (SizeUnit)cboUnit.SelectedIndex;
                KeepRatio = chkKeepRatio.Checked;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                rbPercentage?.Dispose();
                rbFixedSize?.Dispose();
                numPercentage?.Dispose();
                numWidth?.Dispose();
                numHeight?.Dispose();
                cboUnit?.Dispose();
                chkKeepRatio?.Dispose();
                lblPercentageUnit?.Dispose();
                lblWidth?.Dispose();
                lblHeight?.Dispose();
                lblUnit?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
