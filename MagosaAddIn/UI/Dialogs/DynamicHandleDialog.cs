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
    /// 動的調整ハンドル設定用ダイアログ（完全版）
    /// </summary>
    public partial class DynamicHandleDialog : BaseDialog
    {
        public float[] HandleValues { get; private set; }
        public bool DialogResult_OK { get; private set; }

        private List<NumericUpDown> handleControls;
        private Button btnGetCurrentValues;
        private Label lblShapeInfo;
        private GroupBox groupHandles;

        private List<PowerPoint.Shape> targetShapes;
        private ShapeHandleAdjuster adjuster;
        private ShapeHandleAnalysis analysis;

        public DynamicHandleDialog(List<PowerPoint.Shape> shapes, ShapeHandleAnalysis analysis)
        {
            targetShapes = shapes;
            this.analysis = analysis;
            adjuster = new ShapeHandleAdjuster();
            handleControls = new List<NumericUpDown>();

            InitializeComponent();
            CreateDynamicControls();
            UpdateShapeInfo();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int infoHeight = 60;
            const int buttonHeight = 25;
            const int groupWidth = 430;

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

            // 調整ハンドルグループ
            groupHandles = CreateGroupBox("調整ハンドル値（mm単位）", new Point(DefaultMargin, currentY), new Size(groupWidth, 200));

            // 基本コントロールを追加
            this.Controls.AddRange(new Control[] {
                lblShapeInfo,
                btnGetCurrentValues,
                groupHandles
            });

            this.ResumeLayout(false);
        }

        private void CreateDynamicControls()
        {
            // レイアウト定数
            const int formWidth = 490;
            const int groupWidth = 430;
            const int groupInnerMargin = 20;
            const int handleSpacing = 35;
            const int labelWidth = 90;
            const int numericX = 110;
            const int unitX = 220;
            const int descriptionX = 260;
            
            if (analysis == null || analysis.RecommendedHandleCount == 0)
            {
                // 調整ハンドルがない場合
                var lblNoHandles = new Label
                {
                    Text = "選択された図形には調整ハンドルがありません。\n\n" +
                           "調整ハンドル対応図形の例:\n" +
                           "・角丸四角形（角丸の調整）\n" +
                           "・吹き出し（尻尾の位置調整）\n" +
                           "・矢印（矢じりの調整）\n" +
                           "・星形（内側の頂点調整）など",
                    Location = new Point(groupInnerMargin, 30),
                    Size = new Size(groupWidth - 60, 100),
                    ForeColor = Color.Gray
                };
                groupHandles.Controls.Add(lblNoHandles);

                // ボタン位置計算
                int groupBottom = groupHandles.Top + groupHandles.Height;
                int buttonY = groupBottom + ButtonTopMargin;
                int formHeight = CalculateFormHeight(buttonY);

                // フォームサイズを調整
                ConfigureForm("調整ハンドル設定", formWidth, formHeight);

                // ボタンを追加
                AddStandardButtons(buttonY, BtnOK_Click);
                BtnOK.Enabled = false;
                return;
            }

            // 動的にハンドルコントロールを作成
            int handleCount = Math.Min(analysis.RecommendedHandleCount, Constants.MAX_SUPPORTED_HANDLES);
            HandleValues = new float[handleCount];

            // 調整ハンドル図形かどうかを判定
            bool useAdjustmentHandleUnit = analysis.ShapeInfos.Any(info => info.IsAdjustmentHandleShape && !info.IsAngleHandleShape);

            for (int i = 0; i < handleCount; i++)
            {
                int itemY = 30 + i * handleSpacing;
                
                var lblHandle = CreateLabel($"ハンドル {i + 1}:", new Point(groupInnerMargin, itemY), labelWidth);

                var numHandle = CreateNumericUpDown(0, 1, 0, new Point(numericX, itemY - 2), 
                    useAdjustmentHandleUnit ? 2 : 3);
                numHandle.Tag = i; // インデックスを保存

                // 単位に応じて設定を変更
                if (useAdjustmentHandleUnit)
                {
                    // mm単位での設定
                    numHandle.Minimum = (decimal)Constants.MIN_HANDLE_MM;
                    numHandle.Maximum = (decimal)Constants.MAX_HANDLE_MM;
                    numHandle.Value = (decimal)GetInitialHandleValue(i);
                }
                else
                {
                    // 従来の正規化値（0.0-1.0）
                    numHandle.Minimum = (decimal)Constants.MIN_HANDLE_VALUE;
                    numHandle.Maximum = (decimal)Constants.MAX_HANDLE_VALUE;
                    numHandle.Value = (decimal)GetInitialHandleValue(i);
                    numHandle.DecimalPlaces = 3;
                }

                var lblUnit = CreateLabel(useAdjustmentHandleUnit ? "mm" : "", new Point(unitX, itemY), 30);
                lblUnit.ForeColor = Color.Gray;

                var lblDescription = new Label
                {
                    Text = $"調整値 {i + 1}",
                    Location = new Point(descriptionX, itemY),
                    Size = new Size(150, 20),
                    ForeColor = Color.Gray,
                    Font = SmallFont
                };

                handleControls.Add(numHandle);
                groupHandles.Controls.AddRange(new Control[] { lblHandle, numHandle, lblUnit, lblDescription });
            }

            // グループボックスのサイズを調整
            int groupHeight = Math.Max(100, 60 + handleCount * handleSpacing);
            groupHandles.Size = new Size(groupWidth, groupHeight);

            // ボタン位置計算
            int finalGroupBottom = groupHandles.Top + groupHeight;
            int finalButtonY = finalGroupBottom + ButtonTopMargin;
            int finalFormHeight = CalculateFormHeight(finalButtonY);

            // フォームサイズを調整
            ConfigureForm("調整ハンドル設定", formWidth, finalFormHeight);

            // ボタンを追加
            AddStandardButtons(finalButtonY, BtnOK_Click);
        }

        private float GetInitialHandleValue(int handleIndex)
        {
            try
            {
                if (targetShapes != null && targetShapes.Count > 0)
                {
                    var firstShapeWithHandles = targetShapes
                        .FirstOrDefault(s => adjuster.GetHandleInfoFast(s).AdjustmentCount > 0);

                    if (firstShapeWithHandles != null && handleIndex < firstShapeWithHandles.Adjustments.Count)
                    {
                        float currentValue = firstShapeWithHandles.Adjustments[handleIndex + 1]; // PowerPointは1ベース

                        // 調整ハンドル図形かどうかを判定
                        var shapeInfo = adjuster.GetHandleInfoFast(firstShapeWithHandles);
                        bool shouldUseMillimeterUnit = shapeInfo.IsAdjustmentHandleShape && !shapeInfo.IsAngleHandleShape;

                        if (shouldUseMillimeterUnit)
                        {
                            // mm単位の場合
                            float mmValue = adjuster.ConvertNormalizedToMm(currentValue, firstShapeWithHandles, handleIndex);
                            return Math.Max(Constants.MIN_HANDLE_MM, Math.Min(Constants.MAX_HANDLE_MM, mmValue));
                        }
                        else
                        {
                            // 正規化値の場合
                            return Math.Max(Constants.MIN_HANDLE_VALUE, Math.Min(Constants.MAX_HANDLE_VALUE, currentValue));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError($"初期値取得エラー: ハンドル{handleIndex + 1}", ex);
            }

            // デフォルト値を返す
            bool shouldReturnMmDefault = analysis?.ShapeInfos.Any(info => info.IsAdjustmentHandleShape && !info.IsAngleHandleShape) ?? false;
            return shouldReturnMmDefault ? Constants.DEFAULT_HANDLE_MM : Constants.DEFAULT_HANDLE_VALUE;
        }


        private void UpdateShapeInfo()
        {
            if (analysis != null)
            {
                lblShapeInfo.Text = $"選択図形: {analysis.TotalShapes}個\n" +
                                   $"調整ハンドル有り: {analysis.ShapesWithAdjustmentHandles}個\n" +
                                   $"推奨ハンドル数: {analysis.RecommendedHandleCount}個";

                if (analysis.RecommendedHandleCount == 0)
                {
                    lblShapeInfo.Text += "\n※調整ハンドルを持つ図形を選択してください";
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
                    var firstShapeWithHandles = targetShapes
                        .FirstOrDefault(s => adjuster.GetHandleInfoFast(s).AdjustmentCount > 0);

                    if (firstShapeWithHandles != null)
                    {
                        var shapeInfo = adjuster.GetHandleInfoFast(firstShapeWithHandles);
                        bool shouldDisplayInMillimeter = shapeInfo.IsAdjustmentHandleShape && !shapeInfo.IsAngleHandleShape;

                        for (int i = 0; i < Math.Min(handleControls.Count, firstShapeWithHandles.Adjustments.Count); i++)
                        {
                            float currentValue = firstShapeWithHandles.Adjustments[i + 1];

                            if (shouldDisplayInMillimeter)
                            {
                                float mmValue = adjuster.ConvertNormalizedToMm(currentValue, firstShapeWithHandles, i);
                                decimal clampedValue = (decimal)Math.Max(Constants.MIN_HANDLE_MM,
                                    Math.Min(Constants.MAX_HANDLE_MM, mmValue));
                                handleControls[i].Value = clampedValue;
                            }
                            else
                            {
                                decimal clampedValue = (decimal)Math.Max(Constants.MIN_HANDLE_VALUE,
                                    Math.Min(Constants.MAX_HANDLE_VALUE, currentValue));
                                handleControls[i].Value = clampedValue;
                            }
                        }

                        string unitText = shouldDisplayInMillimeter ? "mm単位" : "正規化値";
                        MessageBox.Show($"図形「{firstShapeWithHandles.Name}」の現在値を取得しました。（{unitText}）",
                            "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("調整ハンドルを持つ図形が見つかりません。",
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
                btnGetCurrentValues?.Dispose();
                lblShapeInfo?.Dispose();
                groupHandles?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
