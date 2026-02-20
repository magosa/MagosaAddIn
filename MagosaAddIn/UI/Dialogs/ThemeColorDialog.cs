using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MagosaAddIn.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ColorConv = MagosaAddIn.Core.ColorConverter;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// テーマカラー生成ダイアログ
    /// </summary>
    public partial class ThemeColorDialog : BaseDialog
    {
        #region プロパティ

        public int BaseColor { get; private set; }
        public ColorSchemeType SelectedScheme { get; private set; }
        public int ColorCount { get; private set; }
        public int LightnessSteps { get; private set; }
        public bool ApplyToShapes { get; private set; }
        public bool ArrangePalette { get; private set; }
        public List<int> GeneratedColors { get; private set; }

        #endregion

        #region コントロール

        private GroupBox grpBaseColor;
        private TextBox txtColorCode;
        private Button btnColorPicker;
        private Button btnExtractFromShape;
        private Panel pnlColorPreview;

        private GroupBox grpHueBased;
        private GroupBox grpToneBased;
        private GroupBox grpContrast;

        private RadioButton[] radioHueBased;
        private RadioButton[] radioToneBased;
        private RadioButton[] radioContrast;

        private GroupBox grpOptions;
        private NumericUpDown numColorCount;
        private NumericUpDown numLightnessSteps;

        private GroupBox grpPreview;
        private Panel pnlPreview;
        
        private GroupBox grpActions;
        private CheckBox chkApplyToShapes;
        private CheckBox chkArrangePalette;

        private Button btnApply;
        private Button btnPreview;

        #endregion

        public ThemeColorDialog()
        {
            SetDefaultValues();
            InitializeComponent();
        }

        private void SetDefaultValues()
        {
            BaseColor = 0x5733FF; // デフォルト色（オレンジ系）
            SelectedScheme = ColorSchemeType.Triad;
            ColorCount = 5;
            LightnessSteps = 3;
            ApplyToShapes = true;
            ArrangePalette = false;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            const int formWidth = 540;
            const int margin = 20;
            const int sectionSpacing = 15;
            int currentY = margin;

            // ========== ベースカラー入力 ==========
            grpBaseColor = CreateGroupBox("ベースカラー", new Point(margin, currentY), new Size(formWidth - margin * 2, 80));
            
            var lblColorCode = new Label
            {
                Text = "カラーコード:",
                Location = new Point(15, 28),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            txtColorCode = new TextBox
            {
                Location = new Point(110, 26),
                Size = new Size(90, 22),
                Text = "#FF5733"
            };
            txtColorCode.TextChanged += TxtColorCode_TextChanged;

            btnColorPicker = new Button
            {
                Text = "色選択",
                Location = new Point(210, 24),
                Size = new Size(80, 28)
            };
            btnColorPicker.Click += BtnColorPicker_Click;

            btnExtractFromShape = new Button
            {
                Text = "図形から抽出",
                Location = new Point(300, 24),
                Size = new Size(100, 28)
            };
            btnExtractFromShape.Click += BtnExtractFromShape_Click;

            pnlColorPreview = new Panel
            {
                Location = new Point(410, 24),
                Size = new Size(80, 28),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = ColorConv.RgbToColor(BaseColor)
            };

            grpBaseColor.Controls.AddRange(new Control[] { 
                lblColorCode, txtColorCode, btnColorPicker, btnExtractFromShape, pnlColorPreview 
            });

            currentY += grpBaseColor.Height + sectionSpacing;

            // ========== 配色パターン - 色相ベース ==========
            grpHueBased = CreateGroupBox("色相ベース配色", new Point(margin, currentY), new Size(formWidth - margin * 2, 150));
            CreateHueBasedRadios();
            currentY += grpHueBased.Height + sectionSpacing;

            // ========== 配色パターン - トーンベース ==========
            grpToneBased = CreateGroupBox("トーンベース配色", new Point(margin, currentY), new Size(formWidth - margin * 2, 130));
            CreateToneBasedRadios();
            currentY += grpToneBased.Height + sectionSpacing;

            // ========== 配色パターン - コントラスト ==========
            grpContrast = CreateGroupBox("コントラスト配色", new Point(margin, currentY), new Size(formWidth - margin * 2, 70));
            CreateContrastRadios();
            currentY += grpContrast.Height + sectionSpacing;

            // ========== オプション ==========
            grpOptions = CreateGroupBox("オプション", new Point(margin, currentY), new Size(formWidth - margin * 2, 70));
            
            var lblColorCount = new Label
            {
                Text = "色数:",
                Location = new Point(15, 28),
                Size = new Size(50, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };
            
            numColorCount = CreateNumericUpDown(2, 10, ColorCount, new Point(70, 26));
            numColorCount.ValueChanged += NumericUpDown_ValueChanged;

            var lblLightnessSteps = new Label
            {
                Text = "明度段階:",
                Location = new Point(180, 28),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };
            
            numLightnessSteps = CreateNumericUpDown(
                Constants.MIN_LIGHTNESS_STEPS, 
                Constants.MAX_LIGHTNESS_STEPS, 
                LightnessSteps, 
                new Point(265, 26));
            numLightnessSteps.ValueChanged += NumericUpDown_ValueChanged;

            grpOptions.Controls.AddRange(new Control[] { 
                lblColorCount, numColorCount, lblLightnessSteps, numLightnessSteps 
            });

            currentY += grpOptions.Height + sectionSpacing;

            // ========== プレビュー ==========
            grpPreview = CreateGroupBox("プレビュー", new Point(margin, currentY), new Size(formWidth - margin * 2, 100));
            
            pnlPreview = new Panel
            {
                Location = new Point(15, 25),
                Size = new Size(grpPreview.Width - 30, 60),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            grpPreview.Controls.Add(pnlPreview);

            currentY += grpPreview.Height + sectionSpacing;

            // ========== アクション ==========
            grpActions = CreateGroupBox("実行オプション", new Point(margin, currentY), new Size(formWidth - margin * 2, 70));
            
            chkApplyToShapes = CreateCheckBox("選択図形に適用", new Point(15, 28), new Size(140, 25), true);
            chkArrangePalette = CreateCheckBox("スライド枠外にパレット配置", new Point(170, 28), new Size(200, 25), false);

            grpActions.Controls.AddRange(new Control[] { chkApplyToShapes, chkArrangePalette });

            currentY += grpActions.Height + sectionSpacing + 20;

            // ========== ボタン ==========
            const int buttonTopMargin = 20;
            const int buttonBottomMargin = 60;
            const int buttonHeight = 32;
            const int buttonSpacing = 10;
            
            int buttonY = currentY;

            btnPreview = new Button
            {
                Text = "プレビュー更新",
                Location = new Point(margin, buttonY),
                Size = new Size(130, buttonHeight)
            };
            btnPreview.Click += BtnPreview_Click;

            const int rightButtonWidth = 80;
            int rightButtonX = formWidth - margin - (rightButtonWidth * 2 + buttonSpacing);
            
            btnApply = new Button
            {
                Text = "実行",
                Location = new Point(rightButtonX, buttonY),
                Size = new Size(rightButtonWidth, buttonHeight),
                DialogResult = DialogResult.OK
            };
            btnApply.Click += BtnApply_Click;

            BtnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(rightButtonX + rightButtonWidth + buttonSpacing, buttonY),
                Size = new Size(rightButtonWidth, buttonHeight),
                DialogResult = DialogResult.Cancel
            };

            int formHeight = buttonY + buttonHeight + buttonBottomMargin;
            ConfigureForm("テーマカラー生成", formWidth, formHeight);

            // コントロールを追加
            this.Controls.AddRange(new Control[] {
                grpBaseColor, grpHueBased, grpToneBased, grpContrast,
                grpOptions, grpPreview, grpActions,
                btnPreview, btnApply, BtnCancel
            });

            this.AcceptButton = btnApply;
            this.CancelButton = BtnCancel;

            // 初期プレビュー生成
            UpdatePreview();

            this.ResumeLayout(false);
        }

        #region ラジオボタン作成

        private void CreateHueBasedRadios()
        {
            var schemes = new[]
            {
                (ColorSchemeType.Dyad, "ダイアード（補色）"),
                (ColorSchemeType.Triad, "トライアド（3等分）"),
                (ColorSchemeType.Tetrad, "テトラード（4等分）"),
                (ColorSchemeType.Pentad, "ペンタード（5等分）"),
                (ColorSchemeType.Hexad, "ヘクサード（6等分）"),
                (ColorSchemeType.Analogy, "アナロジー（類似色）"),
                (ColorSchemeType.Intermediate, "インターミディエート（90°）"),
                (ColorSchemeType.Opponent, "オポーネント（135°）"),
                (ColorSchemeType.SplitComplementary, "スプリットコンプリメンタリー")
            };

            radioHueBased = new RadioButton[schemes.Length];
            int baseX = 15;
            int baseY = 25;
            int columnWidth = 175;
            int rowHeight = 25;

            for (int i = 0; i < schemes.Length; i++)
            {
                int col = i / 5;
                int row = i % 5;
                
                radioHueBased[i] = new RadioButton
                {
                    Text = schemes[i].Item2,
                    Location = new Point(baseX + col * columnWidth, baseY + row * rowHeight),
                    Size = new Size(170, 22),
                    Checked = (i == 1), // デフォルトはトライアド
                    Tag = schemes[i].Item1
                };
                radioHueBased[i].CheckedChanged += Radio_CheckedChanged;
                grpHueBased.Controls.Add(radioHueBased[i]);
            }
        }

        private void CreateToneBasedRadios()
        {
            var schemes = new[]
            {
                (ColorSchemeType.ToneOnTone, "トーンオントーン"),
                (ColorSchemeType.ToneInTone, "トーンイントーン"),
                (ColorSchemeType.Camaieu, "カマイユ"),
                (ColorSchemeType.FauxCamaieu, "フォカマイユ"),
                (ColorSchemeType.DominantColor, "ドミナントカラー"),
                (ColorSchemeType.Identity, "アイデンティティ"),
                (ColorSchemeType.Gradation, "グラデーション")
            };

            radioToneBased = new RadioButton[schemes.Length];
            int baseX = 15;
            int baseY = 25;
            int columnWidth = 175;
            int rowHeight = 25;

            for (int i = 0; i < schemes.Length; i++)
            {
                int col = i / 4;
                int row = i % 4;
                
                radioToneBased[i] = new RadioButton
                {
                    Text = schemes[i].Item2,
                    Location = new Point(baseX + col * columnWidth, baseY + row * rowHeight),
                    Size = new Size(170, 22),
                    Tag = schemes[i].Item1
                };
                radioToneBased[i].CheckedChanged += Radio_CheckedChanged;
                grpToneBased.Controls.Add(radioToneBased[i]);
            }
        }

        private void CreateContrastRadios()
        {
            var schemes = new[]
            {
                (ColorSchemeType.HueContrast, "色相コントラスト"),
                (ColorSchemeType.LightnessContrast, "明度コントラスト"),
                (ColorSchemeType.SaturationContrast, "彩度コントラスト")
            };

            radioContrast = new RadioButton[schemes.Length];
            int baseX = 15;
            int spacing = 140;

            for (int i = 0; i < schemes.Length; i++)
            {
                radioContrast[i] = new RadioButton
                {
                    Text = schemes[i].Item2,
                    Location = new Point(baseX + i * spacing, 28),
                    Size = new Size(135, 22),
                    Tag = schemes[i].Item1
                };
                radioContrast[i].CheckedChanged += Radio_CheckedChanged;
                grpContrast.Controls.Add(radioContrast[i]);
            }
        }

        #endregion

        #region イベントハンドラ

        private void TxtColorCode_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string colorCode = txtColorCode.Text.Trim();
                if (colorCode.StartsWith("#") && colorCode.Length == 7)
                {
                    BaseColor = ColorConv.HexToRgb(colorCode);
                    pnlColorPreview.BackColor = ColorConv.RgbToColor(BaseColor);
                }
            }
            catch
            {
                // 無効な色コードは無視
            }
        }

        private void BtnColorPicker_Click(object sender, EventArgs e)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = ColorConv.RgbToColor(BaseColor);
                colorDialog.FullOpen = true;

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    BaseColor = ColorConv.ColorToRgb(colorDialog.Color);
                    txtColorCode.Text = ColorConv.RgbToHex(BaseColor);
                    pnlColorPreview.BackColor = colorDialog.Color;
                    UpdatePreview();
                }
            }
        }

        private void BtnExtractFromShape_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActiveWindow?.Selection != null)
                {
                    var selection = app.ActiveWindow.Selection;
                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && 
                        selection.ShapeRange.Count > 0)
                    {
                        var shape = selection.ShapeRange[1];
                        if (shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            BaseColor = shape.Fill.ForeColor.RGB;
                            txtColorCode.Text = ColorConv.RgbToHex(BaseColor);
                            pnlColorPreview.BackColor = ColorConv.RgbToColor(BaseColor);
                            UpdatePreview();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"図形から色を抽出できませんでした: {ex.Message}", "エラー");
            }
        }

        private void Radio_CheckedChanged(object sender, EventArgs e)
        {
            var radio = sender as RadioButton;
            if (radio != null && radio.Checked && radio.Tag != null)
            {
                SelectedScheme = (ColorSchemeType)radio.Tag;
                UpdatePreview();
            }
        }

        private void NumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void BtnPreview_Click(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            ColorCount = (int)numColorCount.Value;
            LightnessSteps = (int)numLightnessSteps.Value;
            ApplyToShapes = chkApplyToShapes.Checked;
            ArrangePalette = chkArrangePalette.Checked;

            // 最終的な色を生成
            var generator = new ThemeColorGenerator();
            GeneratedColors = generator.GenerateColorScheme(BaseColor, SelectedScheme, ColorCount);
        }

        #endregion

        #region プレビュー生成

        private void UpdatePreview()
        {
            try
            {
                int colorCount = (int)numColorCount.Value;
                var generator = new ThemeColorGenerator();
                var colors = generator.GenerateColorScheme(BaseColor, SelectedScheme, colorCount);

                // プレビューパネルをクリア
                pnlPreview.Controls.Clear();

                // カラーボックスを描画
                int boxWidth = pnlPreview.Width / colors.Count;
                int boxHeight = pnlPreview.Height;

                for (int i = 0; i < colors.Count; i++)
                {
                    var colorBox = new Panel
                    {
                        Location = new Point(i * boxWidth, 0),
                        Size = new Size(boxWidth, boxHeight),
                        BackColor = ColorConv.RgbToColor(colors[i]),
                        BorderStyle = BorderStyle.FixedSingle
                    };
                    pnlPreview.Controls.Add(colorBox);
                }
            }
            catch
            {
                // プレビュー生成失敗は無視
            }
        }

        #endregion

        #region リソース管理

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                grpBaseColor?.Dispose();
                txtColorCode?.Dispose();
                btnColorPicker?.Dispose();
                btnExtractFromShape?.Dispose();
                pnlColorPreview?.Dispose();
                grpHueBased?.Dispose();
                grpToneBased?.Dispose();
                grpContrast?.Dispose();
                grpOptions?.Dispose();
                numColorCount?.Dispose();
                numLightnessSteps?.Dispose();
                grpPreview?.Dispose();
                pnlPreview?.Dispose();
                grpActions?.Dispose();
                chkApplyToShapes?.Dispose();
                chkArrangePalette?.Dispose();
                btnApply?.Dispose();
                btnPreview?.Dispose();

                if (radioHueBased != null)
                    foreach (var radio in radioHueBased) radio?.Dispose();
                if (radioToneBased != null)
                    foreach (var radio in radioToneBased) radio?.Dispose();
                if (radioContrast != null)
                    foreach (var radio in radioContrast) radio?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
