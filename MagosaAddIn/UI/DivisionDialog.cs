using System;
using System.Drawing;
using System.Windows.Forms;

namespace MagosaAddIn.UI
{
    public partial class DivisionDialog : Form
    {
        public int Rows { get; private set; }
        public int Columns { get; private set; }
        public float HorizontalMargin { get; private set; }
        public float VerticalMargin { get; private set; }

        private NumericUpDown numRows;
        private NumericUpDown numColumns;
        private NumericUpDown numHorizontalMargin;
        private NumericUpDown numVerticalMargin;
        private Button btnOK;
        private Button btnCancel;
        private CheckBox chkLinkMargins;

        public DivisionDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Rows = 2;
            Columns = 2;
            HorizontalMargin = 2.0f;
            VerticalMargin = 2.0f;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            this.Text = "図形分割設定";
            this.Size = new Size(350, 280);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 行数設定
            var lblRows = new Label
            {
                Text = "行数:",
                Location = new Point(20, 20),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numRows = new NumericUpDown
            {
                Location = new Point(120, 18),
                Size = new Size(80, 20),
                Minimum = 1,
                Maximum = 50,
                Value = 2
            };

            // 列数設定
            var lblColumns = new Label
            {
                Text = "列数:",
                Location = new Point(20, 50),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numColumns = new NumericUpDown
            {
                Location = new Point(120, 48),
                Size = new Size(80, 20),
                Minimum = 1,
                Maximum = 50,
                Value = 2
            };

            // 水平マージン設定
            var lblHorizontalMargin = new Label
            {
                Text = "水平マージン:",
                Location = new Point(20, 80),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numHorizontalMargin = new NumericUpDown
            {
                Location = new Point(120, 78),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 100,
                Value = 2,
                DecimalPlaces = 1,
                Increment = 0.5m
            };

            var lblHorizontalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 80),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 垂直マージン設定
            var lblVerticalMargin = new Label
            {
                Text = "垂直マージン:",
                Location = new Point(20, 110),
                Size = new Size(90, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numVerticalMargin = new NumericUpDown
            {
                Location = new Point(120, 108),
                Size = new Size(80, 20),
                Minimum = 0,
                Maximum = 100,
                Value = 2,
                DecimalPlaces = 1,
                Increment = 0.5m
            };

            var lblVerticalUnit = new Label
            {
                Text = "pt",
                Location = new Point(210, 110),
                Size = new Size(20, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            //// マージン連動チェックボックス
            //chkLinkMargins = new CheckBox
            //{
            //    Text = "水平・垂直マージンを連動",
            //    Location = new Point(20, 140),
            //    Size = new Size(200, 20),
            //    Checked = true
            //};
            //chkLinkMargins.CheckedChanged += ChkLinkMargins_CheckedChanged;

            // プレビューラベル
            var lblPreview = new Label
            {
                Text = "プレビュー: 2×2 グリッド",
                Location = new Point(20, 170),
                Size = new Size(200, 20),
                ForeColor = Color.Gray
            };

            // 値変更時のプレビュー更新
            numRows.ValueChanged += (s, e) => UpdatePreview(lblPreview);
            numColumns.ValueChanged += (s, e) => UpdatePreview(lblPreview);

            // ボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(120, 210),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(210, 210),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblRows, numRows,
                lblColumns, numColumns,
                lblHorizontalMargin, numHorizontalMargin, lblHorizontalUnit,
                lblVerticalMargin, numVerticalMargin, lblVerticalUnit,
                chkLinkMargins,
                lblPreview,
                btnOK, btnCancel
            });

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            this.ResumeLayout(false);
        }

        private void ChkLinkMargins_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLinkMargins.Checked)
            {
                numVerticalMargin.Value = numHorizontalMargin.Value;
                numVerticalMargin.Enabled = false;
                numHorizontalMargin.ValueChanged += SyncMargins;
            }
            else
            {
                numVerticalMargin.Enabled = true;
                numHorizontalMargin.ValueChanged -= SyncMargins;
            }
        }

        private void SyncMargins(object sender, EventArgs e)
        {
            if (chkLinkMargins.Checked)
            {
                numVerticalMargin.Value = numHorizontalMargin.Value;
            }
        }

        private void UpdatePreview(Label lblPreview)
        {
            lblPreview.Text = $"プレビュー: {numRows.Value}×{numColumns.Value} グリッド";
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Rows = (int)numRows.Value;
            Columns = (int)numColumns.Value;
            HorizontalMargin = (float)numHorizontalMargin.Value;
            VerticalMargin = (float)numVerticalMargin.Value;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numRows?.Dispose();
                numColumns?.Dispose();
                numHorizontalMargin?.Dispose();
                numVerticalMargin?.Dispose();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                chkLinkMargins?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
