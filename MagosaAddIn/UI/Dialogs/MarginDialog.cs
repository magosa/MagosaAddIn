using System;
using System.Drawing;
using System.Windows.Forms;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// マージン設定用ダイアログ
    /// </summary>
    public partial class MarginDialog : BaseDialog
    {
        public new float Margin { get; private set; }

        private NumericUpDown numMargin;

        public MarginDialog(string title)
        {
            InitializeComponent(title);
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Margin = Constants.DEFAULT_HORIZONTAL_MARGIN;
        }

        private void InitializeComponent(string title)
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm(title, 280, 150);

            // マージン設定
            var lblMargin = CreateLabel("マージン:", new Point(DefaultMargin, 20), 60);
            
            numMargin = CreateNumericUpDown((decimal)Constants.MIN_MARGIN, (decimal)Constants.MAX_MARGIN,
                (decimal)Constants.DEFAULT_HORIZONTAL_MARGIN, new Point(90, 18), 2);

            var lblUnit = CreateLabel("pt", new Point(180, 20), 20);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblMargin, numMargin, lblUnit
            });

            // ボタンを追加
            AddStandardButtons(70, BtnOK_Click);

            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Margin = (float)numMargin.Value;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numMargin?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
