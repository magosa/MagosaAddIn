using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 図形置き換え設定用ダイアログ
    /// </summary>
    public partial class ShapeReplacementDialog : BaseDialog
    {
        public SizeMode SelectedSizeMode { get; private set; }
        public bool InheritStyle { get; private set; }
        public bool InheritText { get; private set; }

        private RadioButton rbKeepOriginalSize;
        private RadioButton rbUseTemplateSize;
        private CheckBox chkInheritStyle;
        private CheckBox chkInheritText;
        private Label lblInfo;
        private Label lblNote;

        public ShapeReplacementDialog(int savedShapeCount, string templateShapeName)
        {
            InitializeComponent(savedShapeCount, templateShapeName);
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            SelectedSizeMode = SizeMode.KeepOriginal;
            InheritStyle = false;
            InheritText = false;
        }

        private void InitializeComponent(int savedShapeCount, string templateShapeName)
        {
            this.SuspendLayout();

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int groupWidth = 370;
            const int groupInnerMargin = 20;
            const int radioSpacing = 25;
            const int groupSizeHeight = 80;
            const int groupInheritHeight = 100;

            // 情報表示ラベル
            lblInfo = CreateInfoLabel($"対象図形: {savedShapeCount}個\nテンプレート: {templateShapeName}",
                new Point(DefaultMargin, currentY), new Size(groupWidth, 40));
            currentY += 40 + StandardVerticalSpacing;

            // サイズモード設定グループ
            var groupSize = CreateGroupBox("サイズ設定", new Point(DefaultMargin, currentY), new Size(groupWidth, groupSizeHeight));

            rbKeepOriginalSize = CreateRadioButton("元のサイズを維持",
                new Point(groupInnerMargin, radioSpacing), new Size(330, 20), true);

            rbUseTemplateSize = CreateRadioButton("テンプレートサイズに統一",
                new Point(groupInnerMargin, radioSpacing + 25), new Size(330, 20));

            groupSize.Controls.AddRange(new Control[] {
                rbKeepOriginalSize,
                rbUseTemplateSize
            });
            currentY += groupSizeHeight + StandardVerticalSpacing;

            // スタイル・テキスト継承設定グループ
            var groupInherit = CreateGroupBox("継承設定", new Point(DefaultMargin, currentY), new Size(groupWidth, groupInheritHeight));

            chkInheritStyle = CreateCheckBox("スタイルを継承（塗りつぶし・枠線・影）",
                new Point(groupInnerMargin, radioSpacing), new Size(330, 20));

            chkInheritText = CreateCheckBox("テキストを継承",
                new Point(groupInnerMargin, radioSpacing + 30), new Size(330, 20));

            var lblInheritNote = new Label
            {
                Text = "※チェックなしの場合、テンプレート図形の設定を使用",
                Location = new Point(groupInnerMargin, radioSpacing + 50),
                Size = new Size(330, 20),
                ForeColor = Color.Gray,
                Font = SmallFont
            };

            groupInherit.Controls.AddRange(new Control[] {
                chkInheritStyle,
                chkInheritText,
                lblInheritNote
            });
            currentY += groupInheritHeight + StandardVerticalSpacing;

            // 注意事項ラベル
            lblNote = new Label
            {
                Text = "※ 各図形の中心点の位置は維持されます\n※ 元の図形は削除されます",
                Location = new Point(DefaultMargin, currentY),
                Size = new Size(groupWidth, 35),
                ForeColor = Color.DarkGreen,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9, FontStyle.Italic)
            };
            currentY += 35 + StandardVerticalSpacing;
            
            // ボタンの上にマージンを追加
            currentY += ButtonTopMargin;

            // ボタン位置計算（実行ボタンは高さ28px）
            const int executeButtonHeight = 28;
            int buttonY = currentY;
            int formHeight = CalculateFormHeight(buttonY, executeButtonHeight);

            // フォームの基本設定（幅を広げてボタンマージンを確保）
            ConfigureForm("図形置き換え設定", 440, formHeight);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblInfo,
                groupSize,
                groupInherit,
                lblNote
            });

            // ボタンを追加（カスタムサイズ：幅90px、高さ28px）
            AddStandardButtons(buttonY, BtnOK_Click, 90, executeButtonHeight);
            BtnOK.Text = "実行";
            BtnOK.Font = BoldFont;

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedSizeMode = rbKeepOriginalSize.Checked ? SizeMode.KeepOriginal : SizeMode.UseTemplate;
            InheritStyle = chkInheritStyle.Checked;
            InheritText = chkInheritText.Checked;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                rbKeepOriginalSize?.Dispose();
                rbUseTemplateSize?.Dispose();
                chkInheritStyle?.Dispose();
                chkInheritText?.Dispose();
                lblInfo?.Dispose();
                lblNote?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
