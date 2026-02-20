using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 回転コピー設定用ダイアログ
    /// </summary>
    public partial class RotationCopyDialog : BaseDialog
    {
        public RotationCopyOptions Options { get; private set; }

        private RadioButton rbShapeCenter;
        private RadioButton rbCustomCenter;
        private NumericUpDown numCenterX;
        private NumericUpDown numCenterY;
        private NumericUpDown numAngle;
        private NumericUpDown numCount;
        private Label lblCenterX;
        private Label lblCenterY;
        private Label lblCenterXUnit;
        private Label lblCenterYUnit;

        public RotationCopyDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Options = new RotationCopyOptions
            {
                CenterX = Constants.DEFAULT_CENTER_X,
                CenterY = Constants.DEFAULT_CENTER_Y,
                Angle = Constants.DEFAULT_ROTATION_ANGLE,
                Count = Constants.DEFAULT_ARRAY_COUNT,
                UseShapeCenter = true
            };
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("回転コピー", 320, 300);

            int yPos = InitialTopMargin;

            // 回転中心設定
            var lblCenterMode = CreateLabel("回転中心:", new Point(DefaultMargin, yPos + 2), 80);
            yPos += 25;

            rbShapeCenter = CreateRadioButton("図形の中心", 
                new Point(DefaultMargin + 20, yPos), new Size(150, 20), true);
            rbShapeCenter.CheckedChanged += RbCenterMode_CheckedChanged;
            yPos += 25;

            rbCustomCenter = CreateRadioButton("カスタム座標", 
                new Point(DefaultMargin + 20, yPos), new Size(150, 20), false);
            rbCustomCenter.CheckedChanged += RbCenterMode_CheckedChanged;
            yPos += 30;

            // 中心X座標
            lblCenterX = CreateLabel("X座標:", new Point(DefaultMargin + 40, yPos + 2), 60);
            numCenterX = CreateNumericUpDown((decimal)Constants.MIN_CENTER_COORDINATE, 
                (decimal)Constants.MAX_CENTER_COORDINATE, (decimal)Constants.DEFAULT_CENTER_X,
                new Point(DefaultMargin + 110, yPos), 1);
            lblCenterXUnit = CreateLabel("pt", new Point(DefaultMargin + 200, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 中心Y座標
            lblCenterY = CreateLabel("Y座標:", new Point(DefaultMargin + 40, yPos + 2), 60);
            numCenterY = CreateNumericUpDown((decimal)Constants.MIN_CENTER_COORDINATE, 
                (decimal)Constants.MAX_CENTER_COORDINATE, (decimal)Constants.DEFAULT_CENTER_Y,
                new Point(DefaultMargin + 110, yPos), 1);
            lblCenterYUnit = CreateLabel("pt", new Point(DefaultMargin + 200, yPos + 2), 30);
            yPos += StandardVerticalSpacing + 10;

            // 回転角度
            var lblAngle = CreateLabel("回転角度:", new Point(DefaultMargin, yPos + 2), 80);
            numAngle = CreateNumericUpDown((decimal)Constants.MIN_ROTATION_ANGLE, 
                (decimal)Constants.MAX_ROTATION_ANGLE, (decimal)Constants.DEFAULT_ROTATION_ANGLE,
                new Point(DefaultMargin + 90, yPos), 1);
            var lblAngleUnit = CreateLabel("度", new Point(DefaultMargin + 180, yPos + 2), 30);
            yPos += StandardVerticalSpacing;

            // 個数
            var lblCount = CreateLabel("個数:", new Point(DefaultMargin, yPos + 2), 80);
            numCount = CreateNumericUpDown((decimal)Constants.MIN_ARRAY_COUNT, 
                (decimal)Constants.MAX_ARRAY_COUNT, (decimal)Constants.DEFAULT_ARRAY_COUNT,
                new Point(DefaultMargin + 90, yPos), 0);
            var lblCountUnit = CreateLabel("個", new Point(DefaultMargin + 180, yPos + 2), 30);
            yPos += StandardVerticalSpacing + 10;

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblCenterMode, rbShapeCenter, rbCustomCenter,
                lblCenterX, numCenterX, lblCenterXUnit,
                lblCenterY, numCenterY, lblCenterYUnit,
                lblAngle, numAngle, lblAngleUnit,
                lblCount, numCount, lblCountUnit
            });

            // ボタンを追加
            AddStandardButtons(yPos, BtnOK_Click);

            // フォーム高さを計算して設定
            this.ClientSize = new Size(320, CalculateFormHeight(yPos));

            // 初期状態を設定
            UpdateControlStates();

            this.ResumeLayout(false);
        }

        private void RbCenterMode_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            bool customEnabled = rbCustomCenter.Checked;
            lblCenterX.Enabled = customEnabled;
            lblCenterY.Enabled = customEnabled;
            numCenterX.Enabled = customEnabled;
            numCenterY.Enabled = customEnabled;
            lblCenterXUnit.Enabled = customEnabled;
            lblCenterYUnit.Enabled = customEnabled;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Options = new RotationCopyOptions
            {
                CenterX = (float)numCenterX.Value,
                CenterY = (float)numCenterY.Value,
                Angle = (float)numAngle.Value,
                Count = (int)numCount.Value,
                UseShapeCenter = rbShapeCenter.Checked
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                rbShapeCenter?.Dispose();
                rbCustomCenter?.Dispose();
                numCenterX?.Dispose();
                numCenterY?.Dispose();
                numAngle?.Dispose();
                numCount?.Dispose();
                lblCenterX?.Dispose();
                lblCenterY?.Dispose();
                lblCenterXUnit?.Dispose();
                lblCenterYUnit?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
