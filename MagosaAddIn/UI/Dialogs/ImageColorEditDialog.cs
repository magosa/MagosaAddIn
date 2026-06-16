using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 画像色編集ダイアログ。
    /// 明るさ・コントラスト・色相・彩度・カラーライズ・色調変換（グレースケール/セピア/白黒）を設定する。
    /// </summary>
    public partial class ImageColorEditDialog : BaseDialog
    {
        #region 公開プロパティ

        /// <summary>OK 押下後に取得できる編集設定</summary>
        public Core.ImageColorSettings Settings { get; private set; }

        #endregion

        #region フィールド

        // 明るさ/コントラスト
        private TrackBar _tbBrightness;
        private NumericUpDown _nudBrightness;
        private TrackBar _tbContrast;
        private NumericUpDown _nudContrast;

        // 色相/彩度
        private TrackBar _tbHue;
        private NumericUpDown _nudHue;
        private TrackBar _tbSaturation;
        private NumericUpDown _nudSaturation;

        // カラーライズ
        private CheckBox _chkColorize;
        private TrackBar _tbColorizeIntensity;
        private NumericUpDown _nudColorizeIntensity;
        private Panel _pnlColorPreview;
        private Button _btnColorSelect;

        // 色調変換（排他）
        private RadioButton _rbNone;
        private RadioButton _rbGrayscale;
        private RadioButton _rbSepia;
        private RadioButton _rbBlackWhite;
        private NumericUpDown _nudThreshold;

        // カラーライズ色（PP RGB 形式）
        private int _colorizeRgb = 0xFF; // デフォルト: 赤

        // 更新中フラグ
        private bool _updating = false;

        #endregion

        #region コンストラクタ

        /// <param name="initialShape">（予約）初期値取得用の先頭画像図形。現バージョンでは未使用。</param>
        public ImageColorEditDialog(PowerPoint.Shape initialShape = null)
        {
            InitializeComponent();
        }

        #endregion

        #region UI 初期化

        private void InitializeComponent()
        {
            this.SuspendLayout();

            const int formWidth  = 540;
            const int margin     = DefaultMargin;
            const int grpWidth   = formWidth - margin * 2;
            const int sectionGap = 10;
            int y = margin;

            ConfigureForm("画像色編集", formWidth, 500);

            // ── 明るさ/コントラスト グループ ──────────────────────────────
            var grpBC = CreateGroupBox("明るさ / コントラスト", new Point(margin, y), new Size(grpWidth, 92));
            _tbBrightness  = MakeSlider(-100, 100, 0);
            _nudBrightness = MakeNud(-100, 100, 0);
            _tbContrast    = MakeSlider(-100, 100, 0);
            _nudContrast   = MakeNud(-100, 100, 0);
            LayoutSliderRow(grpBC, "明るさ",      20, _tbBrightness,  _nudBrightness);
            LayoutSliderRow(grpBC, "コントラスト", 52, _tbContrast,    _nudContrast);
            grpBC.Controls.AddRange(new Control[] {
                _tbBrightness, _nudBrightness, _tbContrast, _nudContrast });
            this.Controls.Add(grpBC);
            y += grpBC.Height + sectionGap;

            // ── 色相/彩度 グループ ─────────────────────────────────────────
            var grpHS = CreateGroupBox("色相 / 彩度", new Point(margin, y), new Size(grpWidth, 92));
            _tbHue        = MakeSlider(-180, 180, 0);
            _nudHue       = MakeNud(-180, 180, 0);
            _tbSaturation = MakeSlider(-100, 100, 0);
            _nudSaturation= MakeNud(-100, 100, 0);
            LayoutSliderRow(grpHS, "色相",  20, _tbHue,        _nudHue);
            LayoutSliderRow(grpHS, "彩度",  52, _tbSaturation, _nudSaturation);
            grpHS.Controls.AddRange(new Control[] {
                _tbHue, _nudHue, _tbSaturation, _nudSaturation });
            this.Controls.Add(grpHS);
            y += grpHS.Height + sectionGap;

            // ── カラーライズ グループ ─────────────────────────────────────
            var grpCol = CreateGroupBox("カラーライズ", new Point(margin, y), new Size(grpWidth, 92));

            _chkColorize = CreateCheckBox("カラーライズ", new Point(8, 22), new Size(110, 20));
            _chkColorize.CheckedChanged += ChkColorize_CheckedChanged;

            _tbColorizeIntensity  = MakeSlider(0, 100, 50);
            _nudColorizeIntensity = MakeNud(0, 100, 50);
            LayoutSliderRow(grpCol, "強度", 52, _tbColorizeIntensity, _nudColorizeIntensity,
                enabled: false);

            _pnlColorPreview = new Panel
            {
                Location    = new Point(8, 18),
                Size        = new Size(28, 20),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor   = Color.Red,
                Enabled     = false
            };

            _btnColorSelect = new Button
            {
                Text     = "色選択",
                Location = new Point(8 + 28 + 6, 17),
                Size     = new Size(60, 22),
                Enabled  = false
            };
            _btnColorSelect.Click += BtnColorSelect_Click;

            // カラーライズ行の「強度」ラベルを chkColorize と同じ行に配置し直す
            // chkColorize は左側、強度スライダーを右側に横並びにする
            // 実際の強度スライダーは既に LayoutSliderRow で y=52 に配置済み
            // ここで色選択コントロールを y=18 の右部分に配置
            _pnlColorPreview.Location = new Point(grpWidth - 8 - 60 - 6 - 28, 22);
            _btnColorSelect.Location  = new Point(grpWidth - 8 - 60, 20);

            grpCol.Controls.AddRange(new Control[] {
                _chkColorize, _tbColorizeIntensity, _nudColorizeIntensity,
                _pnlColorPreview, _btnColorSelect });
            this.Controls.Add(grpCol);
            y += grpCol.Height + sectionGap;

            // ── 色調変換 グループ ─────────────────────────────────────────
            var grpTone = CreateGroupBox("色調変換（排他）", new Point(margin, y), new Size(grpWidth, 92));

            _rbNone       = CreateRadioButton("なし",         new Point(10, 22), new Size(55, 20),  true);
            _rbGrayscale  = CreateRadioButton("グレースケール", new Point(70, 22), new Size(120, 20), false);
            _rbSepia      = CreateRadioButton("セピア",        new Point(195, 22), new Size(70, 20), false);
            _rbBlackWhite = CreateRadioButton("白黒",          new Point(10, 52), new Size(60, 20),  false);

            var lblThreshold = new Label
            {
                Text      = "閾値:",
                Location  = new Point(75, 54),
                Size      = new Size(36, 16),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _nudThreshold = MakeNud(0, 100, 50);
            _nudThreshold.Location = new Point(114, 52);
            _nudThreshold.Enabled  = false;

            _rbNone.CheckedChanged       += UpdateThresholdState;
            _rbGrayscale.CheckedChanged  += UpdateThresholdState;
            _rbSepia.CheckedChanged      += UpdateThresholdState;
            _rbBlackWhite.CheckedChanged += UpdateThresholdState;

            grpTone.Controls.AddRange(new Control[] {
                _rbNone, _rbGrayscale, _rbSepia, _rbBlackWhite,
                lblThreshold, _nudThreshold });
            this.Controls.Add(grpTone);
            y += grpTone.Height + sectionGap;

            // ── ボタン行（OK / リセット / キャンセル）──────────────────────
            int btnY = y + ButtonTopMargin;
            int btnRight = formWidth - ButtonRightMargin;

            var btnCancel = new Button
            {
                Text         = "キャンセル",
                Location     = new Point(btnRight - ButtonWidth, btnY),
                Size         = new Size(ButtonWidth, ButtonHeight),
                DialogResult = System.Windows.Forms.DialogResult.Cancel
            };
            var btnOk = new Button
            {
                Text         = "OK",
                Location     = new Point(btnRight - ButtonWidth - ButtonSpacing - ButtonWidth, btnY),
                Size         = new Size(ButtonWidth, ButtonHeight),
                DialogResult = System.Windows.Forms.DialogResult.OK
            };
            btnOk.Click += BtnOK_Click;

            var btnReset = new Button
            {
                Text     = "リセット",
                Location = new Point(margin, btnY),
                Size     = new Size(80, ButtonHeight)
            };
            btnReset.Click += BtnReset_Click;

            this.Controls.AddRange(new Control[] { btnOk, btnCancel, btnReset });
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            this.ClientSize = new Size(formWidth, CalculateFormHeight(btnY));

            // ── TrackBar ↔ NumericUpDown 同期 ──────────────────────────
            WireSync(_tbBrightness,        _nudBrightness);
            WireSync(_tbContrast,          _nudContrast);
            WireSync(_tbHue,               _nudHue);
            WireSync(_tbSaturation,        _nudSaturation);
            WireSync(_tbColorizeIntensity, _nudColorizeIntensity);

            this.ResumeLayout(false);
        }

        // ── コントロールファクトリ ───────────────────────────────────────────

        private TrackBar MakeSlider(int min, int max, int val)
        {
            return new TrackBar
            {
                Minimum    = min,
                Maximum    = max,
                Value      = val,
                TickStyle  = TickStyle.None,
                AutoSize   = false,
                Size       = new Size(295, 20)
            };
        }

        private NumericUpDown MakeNud(int min, int max, int val)
        {
            return new NumericUpDown
            {
                Minimum   = min,
                Maximum   = max,
                Value     = val,
                Increment = 1,
                Size      = new Size(62, 22)
            };
        }

        /// <summary>グループボックス内にスライダー行を配置する</summary>
        private void LayoutSliderRow(GroupBox grp, string label, int rowY,
            TrackBar tb, NumericUpDown nud, bool enabled = true)
        {
            int innerW = grp.Width - 10;
            int lblW   = 78;
            int nudW   = nud.Width;
            int tbW    = innerW - lblW - 5 - nudW - 8;

            grp.Controls.Add(new Label
            {
                Text      = label,
                Location  = new Point(8, rowY + 2),
                Size      = new Size(lblW, 16),
                TextAlign = ContentAlignment.MiddleLeft,
                Enabled   = enabled
            });

            tb.Location = new Point(8 + lblW + 5, rowY);
            tb.Width    = tbW;
            tb.Enabled  = enabled;

            nud.Location = new Point(8 + lblW + 5 + tbW + 5, rowY);
            nud.Enabled  = enabled;
        }

        /// <summary>TrackBar と NumericUpDown を双方向同期させる</summary>
        private void WireSync(TrackBar tb, NumericUpDown nud)
        {
            tb.Scroll += (s, e) =>
            {
                if (_updating) return;
                _updating = true;
                try { if (nud.Value != tb.Value) nud.Value = tb.Value; }
                finally { _updating = false; }
            };
            nud.ValueChanged += (s, e) =>
            {
                if (_updating) return;
                _updating = true;
                try { if (tb.Value != (int)nud.Value) tb.Value = (int)nud.Value; }
                finally { _updating = false; }
            };
        }

        #endregion

        #region イベントハンドラ

        private void ChkColorize_CheckedChanged(object sender, EventArgs e)
        {
            bool on = _chkColorize.Checked;
            _tbColorizeIntensity.Enabled  = on;
            _nudColorizeIntensity.Enabled = on;
            _pnlColorPreview.Enabled      = on;
            _btnColorSelect.Enabled       = on;
        }

        private void UpdateThresholdState(object sender, EventArgs e)
        {
            _nudThreshold.Enabled = _rbBlackWhite.Checked;
        }

        private void BtnColorSelect_Click(object sender, EventArgs e)
        {
            using (var picker = new ColorPickerDialog(_colorizeRgb))
            {
                if (picker.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _colorizeRgb = picker.SelectedColor;
                    int r = _colorizeRgb & 0xFF;
                    int g = (_colorizeRgb >> 8)  & 0xFF;
                    int b = (_colorizeRgb >> 16) & 0xFF;
                    _pnlColorPreview.BackColor = Color.FromArgb(r, g, b);
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Core.ColorToneMode tone;
            if (_rbGrayscale.Checked)    tone = Core.ColorToneMode.Grayscale;
            else if (_rbSepia.Checked)   tone = Core.ColorToneMode.Sepia;
            else if (_rbBlackWhite.Checked) tone = Core.ColorToneMode.BlackAndWhite;
            else                         tone = Core.ColorToneMode.None;

            Settings = new Core.ImageColorSettings
            {
                Brightness          = (int)_nudBrightness.Value,
                Contrast            = (int)_nudContrast.Value,
                Hue                 = (int)_nudHue.Value,
                Saturation          = (int)_nudSaturation.Value,
                ColorizeEnabled     = _chkColorize.Checked,
                ColorizeIntensity   = (int)_nudColorizeIntensity.Value,
                ColorizeRgb         = _colorizeRgb,
                ToneMode            = tone,
                BlackWhiteThreshold = (int)_nudThreshold.Value
            };

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            _updating = true;
            try
            {
                _tbBrightness.Value  = 0; _nudBrightness.Value  = 0;
                _tbContrast.Value    = 0; _nudContrast.Value    = 0;
                _tbHue.Value         = 0; _nudHue.Value         = 0;
                _tbSaturation.Value  = 0; _nudSaturation.Value  = 0;

                _chkColorize.Checked          = false;
                _tbColorizeIntensity.Value    = 50;
                _nudColorizeIntensity.Value   = 50;
                _colorizeRgb                  = 0xFF;
                _pnlColorPreview.BackColor    = Color.Red;

                _rbNone.Checked   = true;
                _nudThreshold.Value = 50;
            }
            finally { _updating = false; }
        }

        #endregion
    }
}
