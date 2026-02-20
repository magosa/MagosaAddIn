using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 円形配列設定用ダイアログ（回転コピー統合版・簡素化版）
    /// </summary>
    public partial class CircularArrayDialog : BaseDialog
    {
        public CircularArrayOptions Options { get; private set; }

        // 配列モード（グループ化）
        private GroupBox grpArrayMode;
        private RadioButton rbEqualDivision;
        private RadioButton rbAngleStep;

        // 回転中心指定（グループ化）
        private GroupBox grpCenterSource;
        private RadioButton rbCustomCoordinate;
        private RadioButton rbTargetShapeCenter;

        // 中心座標
        private Label lblCenterX;
        private Label lblCenterY;
        private NumericUpDown numCenterX;
        private NumericUpDown numCenterY;
        private Label lblCenterXUnit;
        private Label lblCenterYUnit;

        // 等分配置モード用
        private Label lblRadius;
        private NumericUpDown numRadius;
        private Label lblRadiusUnit;

        // 共通
        private Label lblCount;
        private NumericUpDown numCount;
        private Label lblCountUnit;
        private Label lblStartAngle;
        private NumericUpDown numStartAngle;
        private Label lblStartAngleUnit;

        // 角度指定モード用
        private Label lblAngleStep;
        private NumericUpDown numAngleStep;
        private Label lblAngleStepUnit;

        // 図形回転
        private CheckBox chkRotateShapes;

        // 配置設定のベースY位置（動的計算用）
        private int settingsBaseY;

        public CircularArrayDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Options = new CircularArrayOptions
            {
                AngleMode = ArrayAngleMode.EqualDivision,
                CenterSource = CenterSource.CustomCoordinate,
                CenterX = Constants.DEFAULT_CENTER_X,
                CenterY = Constants.DEFAULT_CENTER_Y,
                Radius = Constants.DEFAULT_RADIUS,
                Count = Constants.DEFAULT_ARRAY_COUNT,
                StartAngle = Constants.DEFAULT_START_ANGLE,
                AngleStep = Constants.DEFAULT_ROTATION_ANGLE,
                RotateShapes = false
            };
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("円形配列 / 回転コピー", 360, 450);

            int yPos = InitialTopMargin;

            // ========== 配列モード（GroupBoxで独立グループ化） ==========
            grpArrayMode = new GroupBox
            {
                Text = "配列モード",
                Location = new Point(DefaultMargin, yPos),
                Size = new Size(320, 70)
            };

            rbEqualDivision = new RadioButton
            {
                Text = "等分配置（円周配列）",
                Location = new Point(10, 20),
                Size = new Size(200, 20),
                Checked = true
            };
            rbEqualDivision.CheckedChanged += RbMode_CheckedChanged;

            rbAngleStep = new RadioButton
            {
                Text = "角度指定（回転コピー）",
                Location = new Point(10, 45),
                Size = new Size(200, 20),
                Checked = false
            };
            rbAngleStep.CheckedChanged += RbMode_CheckedChanged;

            grpArrayMode.Controls.Add(rbEqualDivision);
            grpArrayMode.Controls.Add(rbAngleStep);

            yPos += 80;

            // ========== 回転中心指定（GroupBoxで独立グループ化） ==========
            grpCenterSource = new GroupBox
            {
                Text = "回転中心の指定",
                Location = new Point(DefaultMargin, yPos),
                Size = new Size(320, 120)
            };

            rbCustomCoordinate = new RadioButton
            {
                Text = "カスタム座標",
                Location = new Point(10, 20),
                Size = new Size(150, 20),
                Checked = true
            };
            rbCustomCoordinate.CheckedChanged += RbCenterSource_CheckedChanged;

            rbTargetShapeCenter = new RadioButton
            {
                Text = "配列対象図形の中心",
                Location = new Point(10, 45),
                Size = new Size(180, 20),
                Checked = false
            };
            rbTargetShapeCenter.CheckedChanged += RbCenterSource_CheckedChanged;

            grpCenterSource.Controls.Add(rbCustomCoordinate);
            grpCenterSource.Controls.Add(rbTargetShapeCenter);

            // 中心X座標
            lblCenterX = CreateLabel("X座標:", new Point(20, 75), 60);
            numCenterX = CreateNumericUpDown((decimal)Constants.MIN_CENTER_COORDINATE, 
                (decimal)Constants.MAX_CENTER_COORDINATE, (decimal)Constants.DEFAULT_CENTER_X,
                new Point(90, 73), 1);
            lblCenterXUnit = CreateLabel("pt", new Point(180, 75), 30);

            grpCenterSource.Controls.Add(lblCenterX);
            grpCenterSource.Controls.Add(numCenterX);
            grpCenterSource.Controls.Add(lblCenterXUnit);

            // 中心Y座標
            lblCenterY = CreateLabel("Y座標:", new Point(20, 98), 60);
            numCenterY = CreateNumericUpDown((decimal)Constants.MIN_CENTER_COORDINATE, 
                (decimal)Constants.MAX_CENTER_COORDINATE, (decimal)Constants.DEFAULT_CENTER_Y,
                new Point(90, 96), 1);
            lblCenterYUnit = CreateLabel("pt", new Point(180, 98), 30);

            grpCenterSource.Controls.Add(lblCenterY);
            grpCenterSource.Controls.Add(numCenterY);
            grpCenterSource.Controls.Add(lblCenterYUnit);

            yPos += 130;

            // ========== 配置設定 ==========
            var lblSettingsTitle = CreateLabel("配置設定:", new Point(DefaultMargin, yPos), 100);
            lblSettingsTitle.Font = new Font(lblSettingsTitle.Font, FontStyle.Bold);
            yPos += 25;

            // ここから配置設定のベース位置を保存
            settingsBaseY = yPos;

            // 半径（等分配置モード用）
            lblRadius = CreateLabel("半径:", new Point(DefaultMargin, settingsBaseY + 2), 80);
            numRadius = CreateNumericUpDown((decimal)Constants.MIN_RADIUS, 
                (decimal)Constants.MAX_RADIUS, (decimal)Constants.DEFAULT_RADIUS,
                new Point(DefaultMargin + 90, settingsBaseY), 1);
            lblRadiusUnit = CreateLabel("pt", new Point(DefaultMargin + 180, settingsBaseY + 2), 30);

            // 個数
            lblCount = CreateLabel("個数:", new Point(DefaultMargin, 0), 80);
            numCount = CreateNumericUpDown((decimal)Constants.MIN_ARRAY_COUNT, 
                (decimal)Constants.MAX_ARRAY_COUNT, (decimal)Constants.DEFAULT_ARRAY_COUNT,
                new Point(DefaultMargin + 90, 0), 0);
            lblCountUnit = CreateLabel("個", new Point(DefaultMargin + 180, 0), 30);

            // 開始角度
            lblStartAngle = CreateLabel("開始角度:", new Point(DefaultMargin, 0), 80);
            numStartAngle = CreateNumericUpDown((decimal)Constants.MIN_ANGLE_DEGREE, 
                (decimal)Constants.MAX_ANGLE_DEGREE, (decimal)Constants.DEFAULT_START_ANGLE,
                new Point(DefaultMargin + 90, 0), 1);
            lblStartAngleUnit = CreateLabel("度", new Point(DefaultMargin + 180, 0), 30);

            // 角度ステップ（角度指定モード用）
            lblAngleStep = CreateLabel("角度ステップ:", new Point(DefaultMargin, 0), 90);
            numAngleStep = CreateNumericUpDown((decimal)Constants.MIN_ROTATION_ANGLE, 
                (decimal)Constants.MAX_ROTATION_ANGLE, (decimal)Constants.DEFAULT_ROTATION_ANGLE,
                new Point(DefaultMargin + 90, 0), 1);
            lblAngleStepUnit = CreateLabel("度", new Point(DefaultMargin + 180, 0), 30);

            // 図形を回転
            chkRotateShapes = CreateCheckBox("図形を回転させる", 
                new Point(DefaultMargin, 0), new Size(200, 20), false);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                grpArrayMode,
                grpCenterSource,
                lblSettingsTitle,
                lblRadius, numRadius, lblRadiusUnit,
                lblCount, numCount, lblCountUnit,
                lblStartAngle, numStartAngle, lblStartAngleUnit,
                lblAngleStep, numAngleStep, lblAngleStepUnit,
                chkRotateShapes
            });

            // ボタンを追加（仮の位置で作成、後で動的に配置）
            AddStandardButtons(settingsBaseY + 200, BtnOK_Click);

            // 初期状態を設定（位置計算を含む、ボタン位置も更新される）
            UpdateControlStates();

            this.ResumeLayout(false);
        }

        private void RbMode_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void RbCenterSource_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            // 等分配置モードか角度指定モードか
            bool isEqualDivision = rbEqualDivision.Checked;

            // 半径は等分配置モードでのみ表示
            lblRadius.Visible = isEqualDivision;
            numRadius.Visible = isEqualDivision;
            lblRadiusUnit.Visible = isEqualDivision;

            // 角度ステップは角度指定モードでのみ表示
            lblAngleStep.Visible = !isEqualDivision;
            numAngleStep.Visible = !isEqualDivision;
            lblAngleStepUnit.Visible = !isEqualDivision;

            // 中心指定方法による表示制御
            bool isCustomCoordinate = rbCustomCoordinate.Checked;

            lblCenterX.Enabled = isCustomCoordinate;
            lblCenterY.Enabled = isCustomCoordinate;
            numCenterX.Enabled = isCustomCoordinate;
            numCenterY.Enabled = isCustomCoordinate;
            lblCenterXUnit.Enabled = isCustomCoordinate;
            lblCenterYUnit.Enabled = isCustomCoordinate;

            // 動的に配置を再計算
            RepositionSettingsControls();
        }

        /// <summary>
        /// 配置設定のコントロールを動的に再配置
        /// </summary>
        private void RepositionSettingsControls()
        {
            int yPos = settingsBaseY;

            // 等分配置モードの場合は半径を表示
            if (rbEqualDivision.Checked)
            {
                lblRadius.Top = yPos + 2;
                numRadius.Top = yPos;
                lblRadiusUnit.Top = yPos + 2;
                yPos += StandardVerticalSpacing;
            }

            // 個数（共通）
            lblCount.Top = yPos + 2;
            numCount.Top = yPos;
            lblCountUnit.Top = yPos + 2;
            yPos += StandardVerticalSpacing;

            // 開始角度（共通）
            lblStartAngle.Top = yPos + 2;
            numStartAngle.Top = yPos;
            lblStartAngleUnit.Top = yPos + 2;
            yPos += StandardVerticalSpacing;

            // 角度指定モードの場合は角度ステップを表示
            if (!rbEqualDivision.Checked)
            {
                lblAngleStep.Top = yPos + 2;
                numAngleStep.Top = yPos;
                lblAngleStepUnit.Top = yPos + 2;
                yPos += StandardVerticalSpacing;
            }

            // 図形回転チェックボックス
            yPos += 5;
            chkRotateShapes.Top = yPos;
            yPos += StandardVerticalSpacing + 10;

            // ボタンの位置を更新
            UpdateButtonPositions(yPos);

            // フォーム高さを動的に計算
            this.ClientSize = new Size(360, CalculateFormHeight(yPos));
        }

        /// <summary>
        /// OKとキャンセルボタンの位置を更新
        /// </summary>
        private void UpdateButtonPositions(int yPos)
        {
            // BaseDialogのボタンを探して位置を更新
            foreach (Control control in this.Controls)
            {
                if (control is Button btn)
                {
                    if (btn.Text == "OK" || btn.Text == "キャンセル")
                    {
                        btn.Top = yPos;
                    }
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Options = new CircularArrayOptions
            {
                AngleMode = rbEqualDivision.Checked ? ArrayAngleMode.EqualDivision : ArrayAngleMode.AngleStep,
                CenterSource = rbCustomCoordinate.Checked ? CenterSource.CustomCoordinate : CenterSource.TargetShapeCenter,
                CenterX = (float)numCenterX.Value,
                CenterY = (float)numCenterY.Value,
                Radius = (float)numRadius.Value,
                Count = (int)numCount.Value,
                StartAngle = (float)numStartAngle.Value,
                AngleStep = (float)numAngleStep.Value,
                RotateShapes = chkRotateShapes.Checked
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                grpArrayMode?.Dispose();
                grpCenterSource?.Dispose();
                rbEqualDivision?.Dispose();
                rbAngleStep?.Dispose();
                rbCustomCoordinate?.Dispose();
                rbTargetShapeCenter?.Dispose();
                numCenterX?.Dispose();
                numCenterY?.Dispose();
                numRadius?.Dispose();
                numCount?.Dispose();
                numStartAngle?.Dispose();
                numAngleStep?.Dispose();
                chkRotateShapes?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
