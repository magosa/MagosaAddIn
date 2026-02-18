using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// レイヤー（重なり順）調整用ダイアログ
    /// </summary>
    public partial class LayerAdjustmentDialog : BaseDialog
    {
        public LayerOrder SelectedOrder { get; private set; }

        private RadioButton rbSelectionOrderToFront;
        private RadioButton rbSelectionOrderToBack;
        private RadioButton rbLeftToRightToFront;
        private RadioButton rbTopToBottomToFront;
        private Label lblInfo;

        public LayerAdjustmentDialog(int shapeCount)
        {
            InitializeComponent(shapeCount);
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            SelectedOrder = LayerOrder.SelectionOrderToFront;
        }

        private void InitializeComponent(int shapeCount)
        {
            this.SuspendLayout();

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int groupWidth = 400;
            const int groupHeight = 160;
            const int groupInnerMargin = 20;
            const int radioSpacing = 25;
            const int radioVerticalGap = 30;

            // 情報表示ラベル
            lblInfo = CreateInfoLabel($"選択図形: {shapeCount}個\n重なり順を調整します。",
                new Point(DefaultMargin, currentY), new Size(groupWidth, 30));
            currentY += 30 + StandardVerticalSpacing;

            // 調整方法グループ
            var groupOrder = CreateGroupBox("調整方法", new Point(DefaultMargin, currentY), new Size(groupWidth, groupHeight));

            rbSelectionOrderToFront = CreateRadioButton("選択順に前面へ配置（1番目が最背面、最後が最前面）",
                new Point(groupInnerMargin, radioSpacing), new Size(360, 20), true);

            rbSelectionOrderToBack = CreateRadioButton("選択順に背面へ配置（1番目が最前面、最後が最背面）",
                new Point(groupInnerMargin, radioSpacing + radioVerticalGap), new Size(360, 20));

            rbLeftToRightToFront = CreateRadioButton("左から右へ前面に配置（左側が最背面、右側が最前面）",
                new Point(groupInnerMargin, radioSpacing + radioVerticalGap * 2), new Size(360, 20));

            rbTopToBottomToFront = CreateRadioButton("上から下へ前面に配置（上側が最背面、下側が最前面）",
                new Point(groupInnerMargin, radioSpacing + radioVerticalGap * 3), new Size(360, 20));

            groupOrder.Controls.AddRange(new Control[] {
                rbSelectionOrderToFront,
                rbSelectionOrderToBack,
                rbLeftToRightToFront,
                rbTopToBottomToFront
            });
            currentY += groupHeight + StandardVerticalSpacing;
            
            // ボタンの上にマージンを追加
            currentY += ButtonTopMargin;

            // ボタン位置計算（実行ボタンは高さ28px）
            const int executeButtonHeight = 28;
            int buttonY = currentY;
            int formHeight = CalculateFormHeight(buttonY, executeButtonHeight);

            // フォームの基本設定（幅を広げてボタンマージンを確保）
            ConfigureForm("レイヤー（重なり順）調整", 470, formHeight);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblInfo,
                groupOrder
            });

            // ボタンを追加（カスタムサイズ：幅90px、高さ28px）
            AddStandardButtons(buttonY, BtnOK_Click, 90, executeButtonHeight);
            BtnOK.Text = "実行";
            BtnOK.Font = BoldFont;

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (rbSelectionOrderToFront.Checked)
                SelectedOrder = LayerOrder.SelectionOrderToFront;
            else if (rbSelectionOrderToBack.Checked)
                SelectedOrder = LayerOrder.SelectionOrderToBack;
            else if (rbLeftToRightToFront.Checked)
                SelectedOrder = LayerOrder.LeftToRightToFront;
            else if (rbTopToBottomToFront.Checked)
                SelectedOrder = LayerOrder.TopToBottomToFront;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                rbSelectionOrderToFront?.Dispose();
                rbSelectionOrderToBack?.Dispose();
                rbLeftToRightToFront?.Dispose();
                rbTopToBottomToFront?.Dispose();
                lblInfo?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
