using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// HSL スライダーと HEX 入力を持つカスタムカラーピッカーダイアログ。
    /// <para>
    /// SelectedColor は PowerPoint RGB 形式（int 値 0xBBGGRR、R=低バイト）で返す。
    /// </para>
    /// </summary>
    public partial class ColorPickerDialog : BaseDialog
    {
        #region 公開プロパティ

        /// <summary>選択された色（PowerPoint RGB 形式 0xBBGGRR）</summary>
        public int SelectedColor { get; private set; }

        #endregion

        #region フィールド

        // HSL 値（int: H=0-360, S=0-100, L=0-100）
        private int _h = 0;
        private int _s = 100;
        private int _l = 50;

        // 更新中フラグ（ループ防止）
        private bool _updating = false;

        // コントロール
        private GradientBar _hueBar;
        private GradientBar _satBar;
        private GradientBar _lumBar;
        private NumericUpDown _nudH;
        private NumericUpDown _nudS;
        private NumericUpDown _nudL;
        private TextBox _txtHex;
        private Panel _pnlPreview;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// カラーピッカーダイアログを初期化する
        /// </summary>
        /// <param name="initialColor">初期色（PowerPoint RGB 形式 0xBBGGRR）</param>
        public ColorPickerDialog(int initialColor = 0xFF) // デフォルト: 赤
        {
            // PP RGB → HSL に変換して初期スライダー値を設定
            var hsl = MagosaAddIn.Core.ColorConverter.RgbToHsl(initialColor);
            _h = (int)Math.Round(hsl.H);
            _s = (int)Math.Round(hsl.S * 100);
            _l = (int)Math.Round(hsl.L * 100);

            // 範囲クランプ
            _h = Math.Max(0, Math.Min(360, _h));
            _s = Math.Max(0, Math.Min(100, _s));
            _l = Math.Max(0, Math.Min(100, _l));

            SelectedColor = initialColor;

            InitializeComponent();
        }

        #endregion

        #region UI 初期化

        private void InitializeComponent()
        {
            this.SuspendLayout();

            const int formWidth = 420;
            const int margin = 20;
            const int labelW = 22;
            const int barW   = 265;
            const int nudW   = 60;
            const int barH   = 28;
            const int rowH   = 36;
            int y = 15;

            ConfigureForm("色選択", formWidth, 340);

            // ── H スライダー行 ──────────────────────────────────────────────
            AddLabel("H", new Point(margin, y + 8), labelW);
            _hueBar = CreateGradientBar(new Point(margin + labelW + 4, y), barW, barH, 0, 360, _h);
            _hueBar.DrawGradient = DrawHueGradient;
            _nudH = CreateNud(new Point(margin + labelW + 4 + barW + 5, y + 4), nudW, 0, 360, _h);
            this.Controls.Add(_hueBar);
            this.Controls.Add(_nudH);

            y += rowH;

            // ── S スライダー行 ──────────────────────────────────────────────
            AddLabel("S", new Point(margin, y + 8), labelW);
            _satBar = CreateGradientBar(new Point(margin + labelW + 4, y), barW, barH, 0, 100, _s);
            _satBar.DrawGradient = DrawSatGradient;
            _nudS = CreateNud(new Point(margin + labelW + 4 + barW + 5, y + 4), nudW, 0, 100, _s);
            this.Controls.Add(_satBar);
            this.Controls.Add(_nudS);

            y += rowH;

            // ── L スライダー行 ──────────────────────────────────────────────
            AddLabel("L", new Point(margin, y + 8), labelW);
            _lumBar = CreateGradientBar(new Point(margin + labelW + 4, y), barW, barH, 0, 100, _l);
            _lumBar.DrawGradient = DrawLumGradient;
            _nudL = CreateNud(new Point(margin + labelW + 4 + barW + 5, y + 4), nudW, 0, 100, _l);
            this.Controls.Add(_lumBar);
            this.Controls.Add(_nudL);

            y += rowH + 8;

            // ── HEX 入力 ───────────────────────────────────────────────────
            AddLabel("HEX", new Point(margin, y + 4), 38);
            _txtHex = new TextBox
            {
                Location = new Point(margin + 42, y),
                Size     = new Size(120, 22),
                MaxLength = 7,
                Text      = BuildHexString()
            };
            _txtHex.Leave  += TxtHex_Leave;
            _txtHex.KeyDown += TxtHex_KeyDown;
            this.Controls.Add(_txtHex);

            y += 32;

            // ── カラープレビュー ───────────────────────────────────────────
            _pnlPreview = new Panel
            {
                Location  = new Point(margin, y),
                Size      = new Size(formWidth - margin * 2, 50),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = HslToDrawingColor(_h, _s, _l)
            };
            this.Controls.Add(_pnlPreview);

            y += 58;

            // ── OK / キャンセル ────────────────────────────────────────────
            AddStandardButtons(y, BtnOK_Click);
            this.ClientSize = new Size(formWidth, CalculateFormHeight(y));

            // ── イベント接続 ───────────────────────────────────────────────
            _hueBar.ValueChanged += HueBar_ValueChanged;
            _satBar.ValueChanged += SatBar_ValueChanged;
            _lumBar.ValueChanged += LumBar_ValueChanged;
            _nudH.ValueChanged   += NudH_ValueChanged;
            _nudS.ValueChanged   += NudS_ValueChanged;
            _nudL.ValueChanged   += NudL_ValueChanged;

            this.ResumeLayout(false);
        }

        private GradientBar CreateGradientBar(Point loc, int w, int h, int min, int max, int val)
        {
            return new GradientBar
            {
                Location = loc,
                Size     = new Size(w, h),
                Minimum  = min,
                Maximum  = max,
                Value    = val
            };
        }

        private NumericUpDown CreateNud(Point loc, int w, int min, int max, int val)
        {
            return new NumericUpDown
            {
                Location  = loc,
                Size      = new Size(w, 22),
                Minimum   = min,
                Maximum   = max,
                Value     = val,
                Increment = 1
            };
        }

        private void AddLabel(string text, Point loc, int w)
        {
            this.Controls.Add(new Label
            {
                Text      = text,
                Location  = loc,
                Size      = new Size(w, 16),
                TextAlign = ContentAlignment.MiddleRight
            });
        }

        #endregion

        #region グラデーション描画

        // H（色相）スライダー: 全色相レインボー
        private void DrawHueGradient(Graphics g, Rectangle r)
        {
            Color[] stops = { Color.Red, Color.Yellow, Color.Lime, Color.Cyan, Color.Blue, Color.Magenta, Color.Red };
            int segW = r.Width / 6;
            for (int i = 0; i < 6; i++)
            {
                int x1 = r.X + i * segW;
                int x2 = (i == 5) ? r.Right : x1 + segW;
                using (var lgb = new LinearGradientBrush(
                    new Point(x1, r.Y), new Point(x2, r.Y), stops[i], stops[i + 1]))
                {
                    g.FillRectangle(lgb, x1, r.Y, x2 - x1, r.Height);
                }
            }
        }

        // S（彩度）スライダー: グレー → 現在の色相のフル彩度色
        private void DrawSatGradient(Graphics g, Rectangle r)
        {
            Color grey  = HslToDrawingColor(_h, 0, _l);
            Color vivid = HslToDrawingColor(_h, 100, _l);
            using (var lgb = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(r.Right, r.Y), grey, vivid))
            {
                g.FillRectangle(lgb, r);
            }
        }

        // L（明度）スライダー: 黒 → 現在の色 → 白
        private void DrawLumGradient(Graphics g, Rectangle r)
        {
            Color mid   = HslToDrawingColor(_h, _s, 50);
            int   midX  = r.X + r.Width / 2;
            using (var lgb1 = new LinearGradientBrush(
                new Point(r.X, r.Y), new Point(midX, r.Y), Color.Black, mid))
            {
                g.FillRectangle(lgb1, r.X, r.Y, r.Width / 2, r.Height);
            }
            using (var lgb2 = new LinearGradientBrush(
                new Point(midX, r.Y), new Point(r.Right, r.Y), mid, Color.White))
            {
                g.FillRectangle(lgb2, midX, r.Y, r.Width - r.Width / 2, r.Height);
            }
        }

        #endregion

        #region イベントハンドラ

        private void HueBar_ValueChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                _h = _hueBar.Value;
                _nudH.Value = _h;
                _satBar.Invalidate(); // S/L グラデーションは H に依存するため再描画
                _lumBar.Invalidate();
                UpdateFromHSL();
            }
            finally { _updating = false; }
        }

        private void SatBar_ValueChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                _s = _satBar.Value;
                _nudS.Value = _s;
                _lumBar.Invalidate();
                UpdateFromHSL();
            }
            finally { _updating = false; }
        }

        private void LumBar_ValueChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                _l = _lumBar.Value;
                _nudL.Value = _l;
                UpdateFromHSL();
            }
            finally { _updating = false; }
        }

        private void NudH_ValueChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                _h = (int)_nudH.Value;
                _hueBar.Value = _h;
                _satBar.Invalidate();
                _lumBar.Invalidate();
                UpdateFromHSL();
            }
            finally { _updating = false; }
        }

        private void NudS_ValueChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                _s = (int)_nudS.Value;
                _satBar.Value = _s;
                _lumBar.Invalidate();
                UpdateFromHSL();
            }
            finally { _updating = false; }
        }

        private void NudL_ValueChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                _l = (int)_nudL.Value;
                _lumBar.Value = _l;
                UpdateFromHSL();
            }
            finally { _updating = false; }
        }

        private void TxtHex_Leave(object sender, EventArgs e) => ApplyHexInput();
        private void TxtHex_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) ApplyHexInput();
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedColor = MagosaAddIn.Core.ColorConverter.HslToRgb((float)_h, (float)_s / 100f, (float)_l / 100f);
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        #endregion

        #region 内部ヘルパー

        /// <summary>HEX テキストボックスの入力を HSL スライダーに反映する</summary>
        private void ApplyHexInput()
        {
            string hex = _txtHex.Text.Trim().TrimStart('#');
            if (hex.Length != 6) return;

            try
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                // PP RGB 形式に変換
                int ppRgb = r | (g << 8) | (b << 16);

                var hsl2 = MagosaAddIn.Core.ColorConverter.RgbToHsl(ppRgb);
                _h = (int)Math.Round(hsl2.H);
                _s = (int)Math.Round(hsl2.S * 100);
                _l = (int)Math.Round(hsl2.L * 100);
                _h = Math.Max(0, Math.Min(360, _h));
                _s = Math.Max(0, Math.Min(100, _s));
                _l = Math.Max(0, Math.Min(100, _l));

                _updating = true;
                try
                {
                    _hueBar.Value = _h;
                    _satBar.Value = _s;
                    _lumBar.Value = _l;
                    _nudH.Value   = _h;
                    _nudS.Value   = _s;
                    _nudL.Value   = _l;
                    _satBar.Invalidate();
                    _lumBar.Invalidate();
                    UpdatePreview();
                }
                finally { _updating = false; }
            }
            catch { /* 無効な入力は無視 */ }
        }

        /// <summary>HEX 文字列を計算して返す（"#RRGGBB" 形式）</summary>
        private string BuildHexString()
        {
            int ppRgb = MagosaAddIn.Core.ColorConverter.HslToRgb((float)_h, (float)_s / 100f, (float)_l / 100f);
            int r = ppRgb & 0xFF;
            int g = (ppRgb >> 8)  & 0xFF;
            int b = (ppRgb >> 16) & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>HSL から System.Drawing.Color に変換（プレビュー用）</summary>
        private Color HslToDrawingColor(int h, int s, int l)
        {
            int ppRgb = MagosaAddIn.Core.ColorConverter.HslToRgb((float)h, (float)s / 100f, (float)l / 100f);
            int r = ppRgb & 0xFF;
            int g = (ppRgb >> 8)  & 0xFF;
            int b = (ppRgb >> 16) & 0xFF;
            return Color.FromArgb(r, g, b);
        }

        /// <summary>HEX 表示とプレビューを HSL 現在値から更新する</summary>
        private void UpdateFromHSL()
        {
            if (_txtHex != null) _txtHex.Text = BuildHexString();
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (_pnlPreview != null)
                _pnlPreview.BackColor = HslToDrawingColor(_h, _s, _l);
        }

        #endregion
    }
}
