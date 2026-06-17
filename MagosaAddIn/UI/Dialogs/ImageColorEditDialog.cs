using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 画像色編集ダイアログ（改訂版）。
    /// 横長レイアウト：左ペインに操作コントロール、右ペインに変更前/後プレビューを表示する。
    /// </summary>
    public partial class ImageColorEditDialog : BaseDialog
    {
        // ─────────────────────────────────────────────────────────────
        //  レイアウト定数
        // ─────────────────────────────────────────────────────────────
        private const int FormW      = 800;
        private const int LeftW      = 443;  // 左ペイン グループ幅
        private const int RightX     = 463;  // 右ペイン 開始X
        private const int RightW     = 316;  // 右ペイン 幅
        private const int GrpBarH    = 8;    // グラデーションバー高さ
        private const int GrpX       = DefaultMargin;
        private const int SectionGap = 8;

        // ─────────────────────────────────────────────────────────────
        //  公開プロパティ
        // ─────────────────────────────────────────────────────────────
        public Core.ImageColorSettings Settings { get; private set; }

        // ─────────────────────────────────────────────────────────────
        //  左ペイン コントロール
        // ─────────────────────────────────────────────────────────────
        private GradientBar   _gbBrightness, _gbContrast, _gbHue, _gbSaturation, _gbColorizeIntensity;
        private NumericUpDown _nudBrightness, _nudContrast, _nudHue, _nudSaturation, _nudColorizeIntensity;
        private CheckBox      _chkColorize;
        private Panel         _pnlColorPreview;
        private Button        _btnColorSelect;
        private RadioButton   _rbNone, _rbGrayscale, _rbSepia, _rbBlackWhite;
        private GradientBar   _gbThreshold;
        private NumericUpDown _nudThreshold;

        // ─────────────────────────────────────────────────────────────
        //  右ペイン コントロール
        // ─────────────────────────────────────────────────────────────
        private PictureBox _pbBefore, _pbAfter;
        private Label      _lblProcessing;

        // ─────────────────────────────────────────────────────────────
        //  状態
        // ─────────────────────────────────────────────────────────────
        private int          _colorizeRgb = 0xFF;
        private bool         _updating    = false;
        private string       _previewSourcePath;
        private System.Windows.Forms.Timer _debounceTimer;
        private CancellationTokenSource    _previewCts;

        // ─────────────────────────────────────────────────────────────
        //  コンストラクタ
        // ─────────────────────────────────────────────────────────────
        public ImageColorEditDialog(PowerPoint.Shape initialShape = null)
        {
            if (initialShape != null)
            {
                _previewSourcePath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"magosa_prev_src_{Guid.NewGuid():N}.png");
                try   { initialShape.Export(_previewSourcePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG); }
                catch { _previewSourcePath = null; }
            }
            InitializeComponent();
        }

        // ─────────────────────────────────────────────────────────────
        //  フォームイベント
        // ─────────────────────────────────────────────────────────────
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (_previewSourcePath != null && System.IO.File.Exists(_previewSourcePath))
            {
                LoadPictureBox(_pbBefore, _previewSourcePath);
                LoadPictureBox(_pbAfter,  _previewSourcePath);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _previewCts?.Cancel();
            _debounceTimer?.Stop();
            var bi = _pbBefore?.Image; _pbBefore.Image = null; bi?.Dispose();
            var ai = _pbAfter?.Image;  _pbAfter.Image  = null; ai?.Dispose();
            if (_previewSourcePath != null)
                try { System.IO.File.Delete(_previewSourcePath); } catch { }
            base.OnFormClosed(e);
        }

        // ─────────────────────────────────────────────────────────────
        //  UI 初期化
        // ─────────────────────────────────────────────────────────────
        private void InitializeComponent()
        {
            SuspendLayout();
            ConfigureForm("画像色編集", FormW, 100);

            int y = GrpX;

            // ── 明るさ/コントラスト グループ ─────────────────────────
            const int grpH2 = 96;
            var grpBC = CreateGroupBox("明るさ / コントラスト", new Point(GrpX, y), new Size(LeftW, grpH2));
            _gbBrightness  = MakeGradientBar(-100, 100, 0);  _nudBrightness = MakeNud(-100, 100, 0);
            _gbContrast    = MakeGradientBar(-100, 100, 0);  _nudContrast   = MakeNud(-100, 100, 0);
            _gbBrightness.DrawGradient = DrawBrightnessGrad;
            _gbContrast.DrawGradient   = DrawContrastGrad;
            LayoutGradientBarRow(grpBC, "明るさ",       20, _gbBrightness, _nudBrightness);
            LayoutGradientBarRow(grpBC, "コントラスト",  58, _gbContrast,   _nudContrast);
            grpBC.Controls.AddRange(new Control[] {
                _gbBrightness, _nudBrightness,
                _gbContrast,   _nudContrast });
            Controls.Add(grpBC);
            y += grpH2 + SectionGap;

            // ── 色相/彩度 グループ ────────────────────────────────────
            var grpHS = CreateGroupBox("色相 / 彩度", new Point(GrpX, y), new Size(LeftW, grpH2));
            _gbHue        = MakeGradientBar(-180, 180, 0);  _nudHue        = MakeNud(-180, 180, 0);
            _gbSaturation = MakeGradientBar(-100, 100, 0);  _nudSaturation = MakeNud(-100, 100, 0);
            _gbHue.DrawGradient        = DrawHueGrad;
            _gbSaturation.DrawGradient = DrawSatGrad;
            LayoutGradientBarRow(grpHS, "色相", 20, _gbHue,        _nudHue);
            LayoutGradientBarRow(grpHS, "彩度", 58, _gbSaturation, _nudSaturation);
            grpHS.Controls.AddRange(new Control[] {
                _gbHue, _nudHue, _gbSaturation, _nudSaturation });
            Controls.Add(grpHS);
            y += grpH2 + SectionGap;

            // ── カラーライズ グループ ─────────────────────────────────
            const int colGrpH = 88;
            var grpCol = CreateGroupBox("カラーライズ", new Point(GrpX, y), new Size(LeftW, colGrpH));
            _chkColorize = CreateCheckBox("カラーライズ", new Point(8, 22), new Size(110, 20));
            _chkColorize.CheckedChanged += ChkColorize_CheckedChanged;
            _gbColorizeIntensity  = MakeGradientBar(0, 100, 50);
            _nudColorizeIntensity = MakeNud(0, 100, 50);
            _gbColorizeIntensity.DrawGradient = DrawColorizeGrad;
            LayoutGradientBarRow(grpCol, "強度", 46, _gbColorizeIntensity, _nudColorizeIntensity,
                enabled: false);
            _pnlColorPreview = new Panel
            {
                Size = new Size(28, 20), BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Red,   Enabled = false
            };
            _btnColorSelect = new Button { Text = "色選択", Size = new Size(60, 22), Enabled = false };
            _btnColorSelect.Click += BtnColorSelect_Click;
            int ppRight = grpCol.Width - 12;
            _btnColorSelect.Location  = new Point(ppRight - 60, 20);
            _pnlColorPreview.Location = new Point(ppRight - 60 - 6 - 28, 22);
            grpCol.Controls.AddRange(new Control[] {
                _chkColorize,
                _gbColorizeIntensity, _nudColorizeIntensity,
                _pnlColorPreview, _btnColorSelect });
            Controls.Add(grpCol);
            y += colGrpH + SectionGap;

            // ── 色調変換 グループ ─────────────────────────────────────
            const int toneGrpH = 82;
            var grpTone = CreateGroupBox("色調変換（排他）", new Point(GrpX, y), new Size(LeftW, toneGrpH));
            _rbNone       = CreateRadioButton("なし",          new Point(10,  22), new Size(55,  20), true);
            _rbGrayscale  = CreateRadioButton("グレースケール", new Point(70,  22), new Size(120, 20), false);
            _rbSepia      = CreateRadioButton("セピア",         new Point(195, 22), new Size(70,  20), false);
            _rbBlackWhite = CreateRadioButton("白黒",           new Point(10,  48), new Size(60,  20), false);
            var lblThr = new Label { Text = "閾値:", Location = new Point(75, 50),
                Size = new Size(36, 16), TextAlign = ContentAlignment.MiddleLeft };
            _gbThreshold = MakeGradientBar(0, 100, 50);
            _gbThreshold.DrawGradient = DrawThresholdGrad;
            _gbThreshold.Location = new Point(116, 48);
            _gbThreshold.Size     = new Size(240, 28);
            _gbThreshold.Enabled  = false;
            _nudThreshold = MakeNud(0, 100, 50);
            _nudThreshold.Location = new Point(361, 51);
            _nudThreshold.Enabled  = false;
            _rbNone.CheckedChanged       += UpdateThresholdState;
            _rbGrayscale.CheckedChanged  += UpdateThresholdState;
            _rbSepia.CheckedChanged      += UpdateThresholdState;
            _rbBlackWhite.CheckedChanged += UpdateThresholdState;
            // ToneMode が None 以外を選択したらカラーライズを自動解除する
            _rbGrayscale.CheckedChanged  += (s, e) => { if (_rbGrayscale.Checked)  _chkColorize.Checked = false; };
            _rbSepia.CheckedChanged      += (s, e) => { if (_rbSepia.Checked)      _chkColorize.Checked = false; };
            _rbBlackWhite.CheckedChanged += (s, e) => { if (_rbBlackWhite.Checked) _chkColorize.Checked = false; };
            grpTone.Controls.AddRange(new Control[] {
                _rbNone, _rbGrayscale, _rbSepia, _rbBlackWhite, lblThr, _gbThreshold, _nudThreshold });
            Controls.Add(grpTone);
            y += toneGrpH + SectionGap;

            // ── ボタン行 ─────────────────────────────────────────────
            int btnY     = y + ButtonTopMargin;
            int btnRight = GrpX + LeftW - ButtonRightMargin;
            var btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(btnRight - ButtonWidth, btnY),
                Size = new Size(ButtonWidth, ButtonHeight),
                DialogResult = System.Windows.Forms.DialogResult.Cancel
            };
            var btnOk = new Button
            {
                Text = "OK",
                Location = new Point(btnRight - ButtonWidth * 2 - ButtonSpacing, btnY),
                Size = new Size(ButtonWidth, ButtonHeight),
                DialogResult = System.Windows.Forms.DialogResult.OK
            };
            btnOk.Click += BtnOK_Click;
            var btnReset = new Button
            {
                Text     = "リセット",
                Location = new Point(GrpX, btnY),
                Size     = new Size(80, ButtonHeight)
            };
            btnReset.Click += BtnReset_Click;

            Controls.AddRange(new Control[] { btnOk, btnCancel, btnReset });
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            // ── 右ペイン（変更前/後プレビュー）──────────────────────
            int previewW = RightW - 20;
            int previewH = 178;
            int rp       = RightX + 10;

            Controls.Add(new Label { Text = "変更前", Location = new Point(rp, GrpX),
                Size = new Size(60, 18), Font = new Font(Font.FontFamily, 9f, FontStyle.Bold) });
            _pbBefore = new PictureBox
            {
                Location    = new Point(rp, GrpX + 20),
                Size        = new Size(previewW, previewH),
                SizeMode    = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor   = Color.FromArgb(50, 50, 50)
            };
            Controls.Add(_pbBefore);

            int afterLblY = GrpX + 20 + previewH + 10;
            Controls.Add(new Label { Text = "変更後", Location = new Point(rp, afterLblY),
                Size = new Size(60, 18), Font = new Font(Font.FontFamily, 9f, FontStyle.Bold) });
            _pbAfter = new PictureBox
            {
                Location    = new Point(rp, afterLblY + 20),
                Size        = new Size(previewW, previewH),
                SizeMode    = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor   = Color.FromArgb(50, 50, 50)
            };
            Controls.Add(_pbAfter);

            _lblProcessing = new Label
            {
                Text      = "処理中...",
                Location  = new Point(rp, afterLblY + 20),
                Size      = new Size(previewW, previewH),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                Visible   = false
            };
            Controls.Add(_lblProcessing);
            _lblProcessing.BringToFront();

            // ── フォームサイズ確定 ────────────────────────────────────
            int rightBottom = afterLblY + 20 + previewH + GrpX;
            ClientSize = new Size(FormW, Math.Max(CalculateFormHeight(btnY), rightBottom));

            // 色相変更時に彩度バーを再描画
            _gbHue.ValueChanged  += (s, e) => _gbSaturation.Invalidate();
            _nudHue.ValueChanged += (s, e) => _gbSaturation.Invalidate();

            // ── GradientBar ↔ NUD 同期 ──────────────────────────────────
            WireSync(_gbBrightness,        _nudBrightness);
            WireSync(_gbContrast,          _nudContrast);
            WireSync(_gbHue,               _nudHue);
            WireSync(_gbSaturation,        _nudSaturation);
            WireSync(_gbColorizeIntensity, _nudColorizeIntensity);

            // ── debounce タイマー ──────────────────────────────────────
            _debounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _debounceTimer.Tick += (s, e) => { _debounceTimer.Stop(); TriggerPreview(); };

            // プレビュートリガーを全コントロールに登録
            WirePreviewTrigger(_gbBrightness,       _nudBrightness);
            WirePreviewTrigger(_gbContrast,         _nudContrast);
            WirePreviewTrigger(_gbHue,              _nudHue);
            WirePreviewTrigger(_gbSaturation,       _nudSaturation);
            WirePreviewTrigger(_gbColorizeIntensity,_nudColorizeIntensity);
            WireSync(_gbThreshold, _nudThreshold);
            WirePreviewTrigger(_gbThreshold, _nudThreshold);
            _chkColorize.CheckedChanged  += (s, e) => SchedulePreview();
            _rbNone.CheckedChanged       += (s, e) => SchedulePreview();
            _rbGrayscale.CheckedChanged  += (s, e) => SchedulePreview();
            _rbSepia.CheckedChanged      += (s, e) => SchedulePreview();
            _rbBlackWhite.CheckedChanged += (s, e) => SchedulePreview();

            ResumeLayout(false);
        }

        // ─────────────────────────────────────────────────────────────
        //  コントロールファクトリ
        // ─────────────────────────────────────────────────────────────
        private GradientBar MakeGradientBar(int min, int max, int val) =>
            new GradientBar { Minimum = min, Maximum = max, Value = val };

        private NumericUpDown MakeNud(int min, int max, int val) =>
            new NumericUpDown { Minimum = min, Maximum = max, Value = val,
                Increment = 1, Size = new Size(62, 22) };

        private void LayoutGradientBarRow(GroupBox grp, string label, int rowY,
            GradientBar gb, NumericUpDown nud, bool enabled = true)
        {
            int innerW = grp.Width - 10;
            const int lblW = 78;
            int gbW = innerW - lblW - 5 - nud.Width - 5;
            int gbX = 8 + lblW + 5;

            grp.Controls.Add(new Label
            {
                Text = label, Location = new Point(8, rowY + 6),
                Size = new Size(lblW, 16), TextAlign = ContentAlignment.MiddleLeft,
                Enabled = enabled
            });
            gb.Location = new Point(gbX, rowY);
            gb.Size     = new Size(gbW, 28);
            gb.Enabled  = enabled;
            nud.Location = new Point(gbX + gbW + 5, rowY + 3);
            nud.Enabled  = enabled;
        }

        private void WireSync(GradientBar gb, NumericUpDown nud)
        {
            gb.ValueChanged += (s, e) =>
            {
                if (_updating) return;
                _updating = true;
                try { if (nud.Value != gb.Value) nud.Value = gb.Value; }
                finally { _updating = false; }
            };
            nud.ValueChanged += (s, e) =>
            {
                if (_updating) return;
                _updating = true;
                try { if (gb.Value != (int)nud.Value) gb.Value = (int)nud.Value; }
                finally { _updating = false; }
            };
        }

        private void WirePreviewTrigger(GradientBar gb, NumericUpDown nud)
        {
            gb.ValueChanged  += (s, e) => SchedulePreview();
            nud.ValueChanged += (s, e) => SchedulePreview();
        }

        // ─────────────────────────────────────────────────────────────
        //  GradientBar DrawGradient デリゲート
        // ─────────────────────────────────────────────────────────────
        private void DrawBrightnessGrad(Graphics g, Rectangle r)
        {
            if (r.Width <= 1) return;
            using (var lgb = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(r.Right, r.Y), Color.Black, Color.White))
            {
                lgb.InterpolationColors = new ColorBlend(3)
                {
                    Colors    = new[] { Color.Black, Color.FromArgb(128, 128, 128), Color.White },
                    Positions = new[] { 0f, 0.5f, 1f }
                };
                g.FillRectangle(lgb, r);
            }
        }

        private void DrawContrastGrad(Graphics g, Rectangle r)
        {
            if (r.Width <= 1) return;
            using (var lgb = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(r.Right, r.Y),
                Color.FromArgb(96, 96, 96), Color.White))
                g.FillRectangle(lgb, r);
        }

        private void DrawHueGrad(Graphics g, Rectangle r)
        {
            if (r.Width <= 1) return;
            Color[] stops = { Color.Red, Color.Yellow, Color.Lime, Color.Cyan, Color.Blue, Color.Magenta, Color.Red };
            int segW = r.Width / 6;
            for (int i = 0; i < 6; i++)
            {
                int x1 = r.X + i * segW;
                int x2 = (i == 5) ? r.Right : x1 + segW;
                using (var lgb = new LinearGradientBrush(
                    new Point(x1, r.Y), new Point(x2, r.Y), stops[i], stops[i + 1]))
                    g.FillRectangle(lgb, x1, r.Y, x2 - x1, r.Height);
            }
        }

        private void DrawSatGrad(Graphics g, Rectangle r)
        {
            if (r.Width <= 1) return;
            float hDeg     = (_gbHue?.Value ?? 0) + 180f;
            Color grey     = HslToDrawingColor(hDeg, 0f, 0.5f);
            Color saturated = HslToDrawingColor(hDeg, 1f, 0.5f);
            using (var lgb = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(r.Right, r.Y), grey, saturated))
                g.FillRectangle(lgb, r);
        }

        private void DrawThresholdGrad(Graphics g, Rectangle r)
        {
            if (r.Width <= 1) return;
            using (var lgb = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(r.Right, r.Y), Color.Black, Color.White))
                g.FillRectangle(lgb, r);
        }

        private void DrawColorizeGrad(Graphics g, Rectangle r)
        {
            if (r.Width <= 1) return;
            int cr = _colorizeRgb & 0xFF;
            int cg = (_colorizeRgb >> 8)  & 0xFF;
            int cb = (_colorizeRgb >> 16) & 0xFF;
            using (var lgb = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(r.Right, r.Y),
                Color.White, Color.FromArgb(cr, cg, cb)))
                g.FillRectangle(lgb, r);
        }

        // ─────────────────────────────────────────────────────────────
        //  リアルタイムプレビュー
        // ─────────────────────────────────────────────────────────────
        private void SchedulePreview()
        {
            if (_previewSourcePath == null) return;
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void TriggerPreview()
        {
            if (_previewSourcePath == null || IsDisposed) return;
            _previewCts?.Cancel();
            _previewCts = new CancellationTokenSource();
            var ct       = _previewCts.Token;
            var settings = BuildSettings();

            _lblProcessing.Visible = true;
            _lblProcessing.BringToFront();

            Task.Run(() =>
            {
                if (ct.IsCancellationRequested) return;
                string tempOut = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"magosa_prev_out_{Guid.NewGuid():N}.png");
                try
                {
                    new Core.ImageColorEditor().ProcessImage(_previewSourcePath, tempOut, settings);
                    if (ct.IsCancellationRequested) { TryDeleteFile(tempOut); return; }

                    Bitmap bmp;
                    using (var fs = new System.IO.FileStream(
                        tempOut, System.IO.FileMode.Open,
                        System.IO.FileAccess.Read, System.IO.FileShare.Read))
                        bmp = new Bitmap(fs);
                    TryDeleteFile(tempOut);

                    if (ct.IsCancellationRequested) { bmp.Dispose(); return; }
                    if (!IsDisposed)
                        Invoke(new Action(() =>
                        {
                            if (ct.IsCancellationRequested || IsDisposed) { bmp.Dispose(); return; }
                            var old = _pbAfter.Image;
                            _pbAfter.Image = bmp;
                            old?.Dispose();
                            _lblProcessing.Visible = false;
                        }));
                }
                catch
                {
                    TryDeleteFile(tempOut);
                    if (!IsDisposed)
                        try { Invoke(new Action(() => _lblProcessing.Visible = false)); } catch { }
                }
            }, ct);
        }

        private void LoadPictureBox(PictureBox pb, string path)
        {
            try
            {
                using (var fs = new System.IO.FileStream(
                    path, System.IO.FileMode.Open,
                    System.IO.FileAccess.Read, System.IO.FileShare.Read))
                {
                    var bmp = new Bitmap(fs);
                    var old = pb.Image;
                    pb.Image = bmp;
                    old?.Dispose();
                }
            }
            catch { }
        }

        private static void TryDeleteFile(string path)
        {
            try { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); } catch { }
        }

        private static Color HslToDrawingColor(float h, float s, float l)
        {
            int ppRgb = Core.ColorConverter.HslToRgb(h, s, l);
            return Color.FromArgb(ppRgb & 0xFF, (ppRgb >> 8) & 0xFF, (ppRgb >> 16) & 0xFF);
        }

        // ─────────────────────────────────────────────────────────────
        //  イベントハンドラ
        // ─────────────────────────────────────────────────────────────
        private void ChkColorize_CheckedChanged(object sender, EventArgs e)
        {
            bool on = _chkColorize.Checked;
            _gbColorizeIntensity.Enabled  = on;
            _nudColorizeIntensity.Enabled = on;
            _pnlColorPreview.Enabled      = on;
            _btnColorSelect.Enabled       = on;
        }

        private void UpdateThresholdState(object sender, EventArgs e)
        {
            bool bw = _rbBlackWhite.Checked;
            _gbThreshold.Enabled  = bw;
            _nudThreshold.Enabled = bw;
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
                    _gbColorizeIntensity.Invalidate();
                }
            }
        }

        private Core.ImageColorSettings BuildSettings()
        {
            Core.ColorToneMode tone;
            if      (_rbGrayscale.Checked)  tone = Core.ColorToneMode.Grayscale;
            else if (_rbSepia.Checked)      tone = Core.ColorToneMode.Sepia;
            else if (_rbBlackWhite.Checked) tone = Core.ColorToneMode.BlackAndWhite;
            else                            tone = Core.ColorToneMode.None;

            return new Core.ImageColorSettings
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
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Settings = BuildSettings();
            DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            _updating = true;
            try
            {
                _gbBrightness.Value = 0; _nudBrightness.Value = 0;
                _gbContrast.Value   = 0; _nudContrast.Value   = 0;
                _gbHue.Value        = 0; _nudHue.Value        = 0;
                _gbSaturation.Value = 0; _nudSaturation.Value = 0;
                _chkColorize.Checked        = false;
                _gbColorizeIntensity.Value  = 50;
                _nudColorizeIntensity.Value = 50;
                _colorizeRgb               = 0xFF;
                _pnlColorPreview.BackColor  = Color.Red;
                _rbNone.Checked     = true;
                _nudThreshold.Value = 50;
            }
            finally { _updating = false; }
            _gbSaturation.Invalidate();
            _gbColorizeIntensity.Invalidate();
            SchedulePreview();
        }
    }
}
