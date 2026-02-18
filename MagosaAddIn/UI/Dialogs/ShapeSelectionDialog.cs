using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 図形選択条件設定用ダイアログ
    /// </summary>
    public partial class ShapeSelectionDialog : BaseDialog
    {
        public SelectionCriteria SelectedCriteria { get; private set; }
        public int MatchingShapeCount { get; private set; }

        private RadioButton rbFillColorOnly;
        private RadioButton rbLineStyleOnly;
        private RadioButton rbFillAndLineStyle;
        private RadioButton rbShapeTypeOnly;
        private Label lblPreview;
        private Label lblBaseShapeInfo;

        private PowerPoint.Shape baseShape;
        private ShapeSelector selector;

        public ShapeSelectionDialog(PowerPoint.Shape baseShape)
        {
            this.baseShape = baseShape;
            this.selector = new ShapeSelector();
            InitializeComponent();
            SetDefaultValues();
            UpdateBaseShapeInfo();
            UpdatePreview();
        }

        private void SetDefaultValues()
        {
            SelectedCriteria = SelectionCriteria.FillColorOnly;
            MatchingShapeCount = 0;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォームの基本設定
            ConfigureForm("同一書式図形選択", 400, 350);

            // 基準図形情報表示
            lblBaseShapeInfo = CreateInfoLabel("基準図形: ", new Point(DefaultMargin, 20), new Size(350, 40));

            // 選択条件グループボックス
            var groupCriteria = CreateGroupBox("選択条件", new Point(DefaultMargin, 70), new Size(350, 150));

            // 塗りのカラーコードが同じもの
            rbFillColorOnly = CreateRadioButton("塗りのカラーコードが同じもの",
                new Point(20, 25), new Size(300, 20), true);
            rbFillColorOnly.CheckedChanged += RbCriteria_CheckedChanged;

            // 枠線のスタイルが同じもの
            rbLineStyleOnly = CreateRadioButton("枠線のスタイルが同じもの（色・太さ・破線パターン）",
                new Point(20, 50), new Size(300, 20));
            rbLineStyleOnly.CheckedChanged += RbCriteria_CheckedChanged;

            // 塗りと枠線のスタイルが同じもの
            rbFillAndLineStyle = CreateRadioButton("塗りと枠線のスタイルが同じもの",
                new Point(20, 75), new Size(300, 20));
            rbFillAndLineStyle.CheckedChanged += RbCriteria_CheckedChanged;

            // シェイプの種類が同じもの
            rbShapeTypeOnly = CreateRadioButton("シェイプの種類が同じもの（四角形、円、三角形など）",
                new Point(20, 100), new Size(300, 20));
            rbShapeTypeOnly.CheckedChanged += RbCriteria_CheckedChanged;

            groupCriteria.Controls.AddRange(new Control[] {
                rbFillColorOnly, rbLineStyleOnly, rbFillAndLineStyle, rbShapeTypeOnly
            });

            // プレビューラベル
            lblPreview = new Label
            {
                Text = "一致する図形: 0個",
                Location = new Point(DefaultMargin, 230),
                Size = new Size(350, 20),
                ForeColor = Color.Gray,
                Font = ItalicFont
            };

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblBaseShapeInfo,
                groupCriteria,
                lblPreview
            });

            // ボタンを追加
            AddStandardButtons(270, BtnOK_Click);
            BtnOK.Text = "選択実行";
            BtnOK.Size = new Size(80, 25);

            this.ResumeLayout(false);
        }

        private void UpdateBaseShapeInfo()
        {
            if (baseShape != null)
            {
                try
                {
                    var shapeName = ComExceptionHandler.ExecuteComOperation(
                        () => baseShape.Name,
                        "基準図形名取得",
                        defaultValue: "不明な図形",
                        suppressErrors: true);

                    lblBaseShapeInfo.Text = $"基準図形: {shapeName}";
                }
                catch
                {
                    lblBaseShapeInfo.Text = "基準図形: 情報取得エラー";
                }
            }
        }

        private void RbCriteria_CheckedChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            try
            {
                var criteria = GetSelectedCriteria();
                var count = selector.GetMatchingShapeCount(baseShape, criteria);
                MatchingShapeCount = count;

                lblPreview.Text = $"一致する図形: {count}個";
                lblPreview.ForeColor = count > 0 ? Color.DarkGreen : Color.Gray;

                // OKボタンの有効/無効制御
                BtnOK.Enabled = count > 0;
            }
            catch (Exception ex)
            {
                lblPreview.Text = "プレビュー取得エラー";
                lblPreview.ForeColor = Color.Red;
                BtnOK.Enabled = false;
                ComExceptionHandler.LogError("選択プレビュー更新", ex);
            }
        }

        private SelectionCriteria GetSelectedCriteria()
        {
            if (rbFillColorOnly.Checked)
                return SelectionCriteria.FillColorOnly;
            else if (rbLineStyleOnly.Checked)
                return SelectionCriteria.LineStyleOnly;
            else if (rbFillAndLineStyle.Checked)
                return SelectionCriteria.FillAndLineStyle;
            else if (rbShapeTypeOnly.Checked)
                return SelectionCriteria.ShapeTypeOnly;
            else
                return SelectionCriteria.FillColorOnly;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedCriteria = GetSelectedCriteria();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                rbFillColorOnly?.Dispose();
                rbLineStyleOnly?.Dispose();
                rbFillAndLineStyle?.Dispose();
                rbShapeTypeOnly?.Dispose();
                lblPreview?.Dispose();
                lblBaseShapeInfo?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
