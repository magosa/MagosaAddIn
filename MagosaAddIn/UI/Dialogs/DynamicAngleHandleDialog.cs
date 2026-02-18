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
    /// 動的角度ハンドル設定用ダイアログ
    /// </summary>
    public partial class DynamicAngleHandleDialog : BaseDialog
    {
        public float[] HandleValues { get; private set; }
        public bool DialogResult_OK { get; private set; }

        private List<NumericUpDown> handleControls;
        private List<Label> interpretationLabels;
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

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int infoHeight = 80;
            const int buttonHeight = 25;
            const int groupWidth = 480;

            // 図形情報表示
            lblShapeInfo = CreateInfoLabel("図形情報を取得中...", new Point(DefaultMargin, currentY), new Size(groupWidth, infoHeight));
            currentY += infoHeight + StandardVerticalSpacing;

            // 現在値取得ボタン
            btnGetCurrentValues = new Button
            {
                Text = "現在の値を取得",
                Location = new Point(DefaultMargin, currentY),
                Size = new Size(120, buttonHeight)
            };
            btnGetCurrentValues.Click += BtnGetCurrentValues_Click;
            currentY += buttonHeight + StandardVerticalSpacing;

            // 角度ハンドルグループ
            groupAngleHandles = CreateGroupBox("角度ハンドル値（度数）", new Point(DefaultMargin, currentY), new Size(groupWidth, 200));

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
            // レイアウト定数
            const int formWidth = 540;
            const int groupWidth = 480;
            const int groupInnerMargin = 20;
            const int handleSpacing = 50;
            const int labelWidth = 100;
            const int numericX = 130;
            const int unitX = 240;
            const int descriptionX = 270;
            
            if (analysis == null || analysis.RecommendedAngleHandleCount == 0)
            {
                // 角度ハンドルがない場合
                var lblNoHandles = new Label
                {
                    Text = "選択された図形には角度ハンドルがありません。\n" +
                           "角度ハンドル対応図形: 円弧、弦、扇形、ブロック円弧、ドーナツ、三日月など",
                    Location = new Point(groupInnerMargin, 30),
                    Size = new Size(groupWidth - 60, 40),
                    ForeColor = Color.Gray
                };
                groupAngleHandles.Controls.Add(lblNoHandles);

                // ボタン位置計算
                int groupBottom = groupAngleHandles.Top + groupAngleHandles.Height;
                int buttonY = groupBottom + ButtonTopMargin;
                int formHeight = CalculateFormHeight(buttonY);

                // フォームサイズを調整
                ConfigureForm("角度ハンドル設定", formWidth, formHeight);

                // ボタンを追加
                AddStandardButtons(buttonY, BtnOK_Click);
                BtnOK.Enabled = false;
                return;
            }

            // 角度ハンドル図形の代表例を取得
            var representativeShape = analysis.ShapeInfos
                .FirstOrDefault(info => info.IsAngleHandleShape && info.AdjustmentCount > 0);

            if (representativeShape == null)
            {
                CreateDynamicAngleControls(); // 再帰的に呼び出し
                return;
            }

            // 動的に角度ハンドルコントロールを作成
            int handleCount = Math.Min(representativeShape.AdjustmentCount, Constants.MAX_SUPPORTED_HANDLES);
            HandleValues = new float[handleCount];

            for (int i = 0; i < handleCount; i++)
            {
                int itemY = 30 + i * handleSpacing;
                
                // ハンドルの意味を表示
                string handleMeaning = i < representativeShape.AngleInterpretation.Count
                    ? representativeShape.AngleInterpretation[i]
                    : $"ハンドル{i + 1}";

                var lblHandle = CreateLabel($"{handleMeaning}:", new Point(groupInnerMargin, itemY), labelWidth);
                lblHandle.Font = BoldFont;

                var numHandle = CreateNumericUpDown((decimal)Constants.MIN_ANGLE_DEGREE, (decimal)Constants.MAX_ANGLE_DEGREE,
                    (decimal)GetInitialAngleValue(i), new Point(numericX, itemY - 2), 2);
                numHandle.Tag = i; // インデックスを保存

                var lblUnit = CreateLabel("°", new Point(unitX, itemY), 20);
                lblUnit.ForeColor = Color.Gray;

                // 角度の説明ラベル
                var lblInterpretation = new Label
                {
                    Text = GetAngleDescription(representativeShape.ShapeType, i),
                    Location = new Point(descriptionX, itemY),
                    Size = new Size(200, 20),
                    ForeColor = Color.Gray,
                    Font = SmallFont
                };

                handleControls.Add(numHandle);
                interpretationLabels.Add(lblInterpretation);
                groupAngleHandles.Controls.AddRange(new Control[] { lblHandle, numHandle, lblUnit, lblInterpretation });
            }

            // グループボックスのサイズを調整
            int groupHeight = Math.Max(100, 60 + handleCount * handleSpacing);
            groupAngleHandles.Size = new Size(groupWidth, groupHeight);

            // ボタン位置計算
            int finalGroupBottom = groupAngleHandles.Top + groupHeight;
            int finalButtonY = finalGroupBottom + ButtonTopMargin;
            int finalFormHeight = CalculateFormHeight(finalButtonY);

            // フォームサイズを調整
            ConfigureForm("角度ハンドル設定", formWidth, finalFormHeight);

            // ボタンを追加
            AddStandardButtons(finalButtonY, BtnOK_Click);
        }

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
                btnGetCurrentValues?.Dispose();
                lblShapeInfo?.Dispose();
                groupAngleHandles?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
