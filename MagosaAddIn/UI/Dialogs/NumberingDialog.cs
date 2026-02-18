using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 自動ナンバリング用ダイアログ
    /// </summary>
    public partial class NumberingDialog : BaseDialog
    {
        public int StartNumber { get; private set; }
        public int Increment { get; private set; }
        public NumberFormat SelectedFormat { get; private set; }
        public float FontSize { get; private set; }

        private NumericUpDown numStartNumber;
        private NumericUpDown numIncrement;
        private NumericUpDown numFontSize;
        private ComboBox cmbFormat;
        private Label lblPreview;
        private Label lblInfo;

        public NumberingDialog(int shapeCount)
        {
            InitializeComponent(shapeCount);
            SetDefaultValues();
            UpdatePreview();
        }

        private void SetDefaultValues()
        {
            StartNumber = Constants.DEFAULT_START_NUMBER;
            Increment = Constants.DEFAULT_INCREMENT;
            SelectedFormat = NumberFormat.Arabic;
            FontSize = Constants.DEFAULT_NUMBER_FONT_SIZE;
        }

        private void InitializeComponent(int shapeCount)
        {
            this.SuspendLayout();

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int labelWidth = 100;
            const int numericX = 130;
            const int comboWidth = 200;
            const int previewHeight = 80;

            // 情報表示ラベル
            lblInfo = CreateInfoLabel($"選択図形: {shapeCount}個\n選択順に番号を付けます。",
                new Point(DefaultMargin, currentY), new Size(400, 30));
            currentY += 30 + StandardVerticalSpacing;

            // 開始番号
            var lblStartNumber = CreateLabel("開始番号:", new Point(DefaultMargin, currentY), labelWidth);
            numStartNumber = CreateNumericUpDown(Constants.MIN_START_NUMBER, Constants.MAX_START_NUMBER,
                Constants.DEFAULT_START_NUMBER, new Point(numericX, currentY - 2));
            currentY += StandardVerticalSpacing;

            // 増分値
            var lblIncrement = CreateLabel("増分値:", new Point(DefaultMargin, currentY), labelWidth);
            numIncrement = CreateNumericUpDown(Constants.MIN_INCREMENT, Constants.MAX_INCREMENT,
                Constants.DEFAULT_INCREMENT, new Point(numericX, currentY - 2));
            numIncrement.ValueChanged += (s, e) => UpdatePreview();
            currentY += StandardVerticalSpacing;

            // 番号フォーマット
            var lblFormat = CreateLabel("番号フォーマット:", new Point(DefaultMargin, currentY), labelWidth);
            cmbFormat = CreateComboBox(new Point(numericX, currentY - 2), new Size(comboWidth, 20));
            cmbFormat.Items.AddRange(new object[] {
                "1, 2, 3... (算用数字)",
                "①②③... (丸数字)",
                "A, B, C... (大文字)",
                "a, b, c... (小文字)",
                "I, II, III... (ローマ数字大文字)",
                "i, ii, iii... (ローマ数字小文字)"
            });
            cmbFormat.SelectedIndex = 0;
            cmbFormat.SelectedIndexChanged += (s, e) => UpdatePreview();
            currentY += StandardVerticalSpacing;

            // フォントサイズ
            var lblFontSize = CreateLabel("フォントサイズ:", new Point(DefaultMargin, currentY), labelWidth);
            numFontSize = CreateNumericUpDown(8, 72, (decimal)Constants.DEFAULT_NUMBER_FONT_SIZE,
                new Point(numericX, currentY - 2));

            var lblUnit = new Label
            {
                Text = "pt",
                Location = new Point(numericX + NumericUpDownWidth + 10, currentY),
                Size = new Size(20, 20),
                ForeColor = Color.Gray
            };
            currentY += StandardVerticalSpacing + 10;

            // プレビュー
            var lblPreviewTitle = CreateLabel("プレビュー:", new Point(DefaultMargin, currentY), labelWidth);
            lblPreviewTitle.Font = BoldFont;
            currentY += 25;

            lblPreview = new Label
            {
                Location = new Point(DefaultMargin, currentY),
                Size = new Size(400, previewHeight),
                Text = "",
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 12)
            };
            currentY += previewHeight + StandardVerticalSpacing;
            
            // ボタンの上にマージンを追加
            currentY += ButtonTopMargin;

            // ボタン位置計算（適用ボタンは高さ28px）
            const int executeButtonHeight = 28;
            int buttonY = currentY;
            int formHeight = CalculateFormHeight(buttonY, executeButtonHeight);

            // フォームの基本設定（幅を広げてボタンマージンを確保）
            ConfigureForm("自動ナンバリング", 470, formHeight);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblInfo,
                lblStartNumber, numStartNumber,
                lblIncrement, numIncrement,
                lblFormat, cmbFormat,
                lblFontSize, numFontSize, lblUnit,
                lblPreviewTitle, lblPreview
            });

            // ボタンを追加（カスタムサイズ：幅90px、高さ28px）
            AddStandardButtons(buttonY, BtnOK_Click, 90, executeButtonHeight);
            BtnOK.Text = "適用";
            BtnOK.Font = BoldFont;

            this.ResumeLayout(false);
        }

        private void UpdatePreview()
        {
            try
            {
                int start = (int)numStartNumber.Value;
                int inc = (int)numIncrement.Value;
                NumberFormat format = (NumberFormat)cmbFormat.SelectedIndex;

                var previewNumbers = new List<string>();

                for (int i = 0; i < Math.Min(5, 10); i++)
                {
                    int num = start + (i * inc);
                    string formatted = FormatNumberForPreview(num, format);
                    previewNumbers.Add(formatted);
                }

                lblPreview.Text = string.Join(", ", previewNumbers) + "...";
            }
            catch
            {
                lblPreview.Text = "プレビュー取得エラー";
            }
        }

        private string FormatNumberForPreview(int number, NumberFormat format)
        {
            switch (format)
            {
                case NumberFormat.Arabic:
                    return number.ToString();
                case NumberFormat.CircledArabic:
                    if (number >= 1 && number <= 20)
                        return char.ConvertFromUtf32(0x245F + number);
                    return $"({number})";
                case NumberFormat.UpperAlpha:
                    return GetAlpha(number, true);
                case NumberFormat.LowerAlpha:
                    return GetAlpha(number, false);
                case NumberFormat.UpperRoman:
                    return GetRoman(number, true);
                case NumberFormat.LowerRoman:
                    return GetRoman(number, false);
                default:
                    return number.ToString();
            }
        }

        private string GetAlpha(int number, bool isUpper)
        {
            if (number < 1) return isUpper ? "A" : "a";
            string result = "";
            int n = number;
            while (n > 0)
            {
                n--;
                result = (char)((isUpper ? 'A' : 'a') + (n % 26)) + result;
                n = n / 26;
            }
            return result;
        }

        private string GetRoman(int number, bool isUpper)
        {
            if (number < 1 || number > 3999) return number.ToString();
            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] upper = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            string[] lower = { "m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i" };
            string[] romans = isUpper ? upper : lower;
            string result = "";
            for (int i = 0; i < values.Length; i++)
            {
                while (number >= values[i])
                {
                    number -= values[i];
                    result += romans[i];
                }
            }
            return result;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            StartNumber = (int)numStartNumber.Value;
            Increment = (int)numIncrement.Value;
            SelectedFormat = (NumberFormat)cmbFormat.SelectedIndex;
            FontSize = (float)numFontSize.Value;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numStartNumber?.Dispose();
                numIncrement?.Dispose();
                numFontSize?.Dispose();
                cmbFormat?.Dispose();
                lblPreview?.Dispose();
                lblInfo?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
