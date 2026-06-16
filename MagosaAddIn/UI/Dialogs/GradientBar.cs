using System;
using System.Drawing;
using System.Windows.Forms;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// グラデーション背景に白いサムを持つカスタムスライダーコントロール。
    /// ColorPickerDialog および ImageColorEditDialog で共有して使用する。
    /// </summary>
    internal class GradientBar : Control
    {
        public event EventHandler ValueChanged;

        private int _minimum = 0;
        private int _maximum = 100;
        private int _value   = 50;

        public int Minimum { get => _minimum; set { _minimum = value; Invalidate(); } }
        public int Maximum { get => _maximum; set { _maximum = value; Invalidate(); } }

        public int Value
        {
            get => _value;
            set
            {
                int clamped = Math.Max(_minimum, Math.Min(_maximum, value));
                if (clamped == _value) return;
                _value = clamped;
                Invalidate();
                ValueChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        /// <summary>グラデーション描画デリゲート。null の場合は無地表示。</summary>
        public Action<Graphics, Rectangle> DrawGradient { get; set; }

        public GradientBar()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.ResizeRedraw, true);
            Cursor = Cursors.Hand;
            Size   = new Size(200, 28);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            // グラデーション帯（上下マージン 6px）
            var gradRect = new Rectangle(6, 6, Width - 12, Height - 12);

            if (DrawGradient != null)
                DrawGradient(e.Graphics, gradRect);
            else
                e.Graphics.FillRectangle(SystemBrushes.ControlLight, gradRect);

            e.Graphics.DrawRectangle(Pens.DarkGray, gradRect);

            // サム（縦線）
            if (_maximum > _minimum)
            {
                double ratio  = (double)(_value - _minimum) / (_maximum - _minimum);
                int    thumbX = gradRect.X + (int)(ratio * gradRect.Width);
                thumbX = Math.Max(gradRect.X, Math.Min(gradRect.Right, thumbX));

                using (var wp = new Pen(Color.White, 3))
                    e.Graphics.DrawLine(wp, thumbX, 2, thumbX, Height - 3);
                using (var bp = new Pen(Color.Black, 1))
                    e.Graphics.DrawLine(bp, thumbX, 2, thumbX, Height - 3);
            }
        }

        protected override void OnMouseDown(MouseEventArgs e) => SetFromMouse(e.X);
        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) SetFromMouse(e.X);
        }

        private void SetFromMouse(int x)
        {
            var gradRect = new Rectangle(6, 6, Width - 12, Height - 12);
            if (gradRect.Width <= 0) return;
            double ratio = Math.Max(0, Math.Min(1, (double)(x - gradRect.X) / gradRect.Width));
            Value = _minimum + (int)Math.Round(ratio * (_maximum - _minimum));
        }
    }
}
