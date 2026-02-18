using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;
using System.Collections.Generic;
using System.Linq;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 円形配置用ダイアログ
    /// </summary>
    public partial class CircleArrangementDialog : BaseDialog
    {
        public float CenterX { get; private set; }
        public float CenterY { get; private set; }
        public float Radius { get; private set; }

        private NumericUpDown numCenterX;
        private NumericUpDown numCenterY;
        private NumericUpDown numRadius;
        private Button btnUseCurrentCenter;

        public CircleArrangementDialog()
        {
            InitializeComponent();
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            CenterX = Constants.DEFAULT_CENTER_X;
            CenterY = Constants.DEFAULT_CENTER_Y;
            Radius = Constants.DEFAULT_RADIUS;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // レイアウト計算
            int currentY = InitialTopMargin;
            const int numericX = 120;
            const int buttonHeight = 25;

            // 中心X座標設定
            var lblCenterX = CreateLabel("中心X座標:", new Point(DefaultMargin, currentY));
            numCenterX = CreateNumericUpDown((decimal)Constants.MIN_CENTER_COORDINATE, (decimal)Constants.MAX_CENTER_COORDINATE,
                (decimal)Constants.DEFAULT_CENTER_X, new Point(numericX, currentY - 2), 2);
            currentY += StandardVerticalSpacing;

            // 中心Y座標設定
            var lblCenterY = CreateLabel("中心Y座標:", new Point(DefaultMargin, currentY));
            numCenterY = CreateNumericUpDown((decimal)Constants.MIN_CENTER_COORDINATE, (decimal)Constants.MAX_CENTER_COORDINATE,
                (decimal)Constants.DEFAULT_CENTER_Y, new Point(numericX, currentY - 2), 2);
            currentY += StandardVerticalSpacing;

            // 半径設定
            var lblRadius = CreateLabel("半径:", new Point(DefaultMargin, currentY));
            numRadius = CreateNumericUpDown((decimal)Constants.MIN_RADIUS, (decimal)Constants.MAX_RADIUS,
                (decimal)Constants.DEFAULT_RADIUS, new Point(numericX, currentY - 2), 2);
            currentY += StandardVerticalSpacing;

            // 現在の中心を使用ボタン
            btnUseCurrentCenter = new Button
            {
                Text = "選択図形の中心を使用",
                Location = new Point(DefaultMargin, currentY),
                Size = new Size(150, buttonHeight)
            };
            btnUseCurrentCenter.Click += BtnUseCurrentCenter_Click;
            currentY += buttonHeight + StandardVerticalSpacing;
            
            // ボタンの上にマージンを追加
            currentY += ButtonTopMargin;

            // ボタン位置計算
            int buttonY = currentY;
            int formHeight = CalculateFormHeight(buttonY);

            // フォームの基本設定
            ConfigureForm("円形配置設定", 300, formHeight);

            // コントロールをフォームに追加
            this.Controls.AddRange(new Control[] {
                lblCenterX, numCenterX,
                lblCenterY, numCenterY,
                lblRadius, numRadius,
                btnUseCurrentCenter
            });

            // ボタンを追加
            AddStandardButtons(buttonY, BtnOK_Click);

            this.ResumeLayout(false);
        }

        private void BtnUseCurrentCenter_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActiveWindow?.Selection != null)
                {
                    var selection = app.ActiveWindow.Selection;
                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        var shapes = new List<PowerPoint.Shape>();
                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            shapes.Add(selection.ShapeRange[i]);
                        }

                        if (shapes.Count > 0)
                        {
                            float minLeft = shapes.Min(s => s.Left);
                            float maxRight = shapes.Max(s => s.Left + s.Width);
                            float minTop = shapes.Min(s => s.Top);
                            float maxBottom = shapes.Max(s => s.Top + s.Height);

                            numCenterX.Value = (decimal)((minLeft + maxRight) / 2);
                            numCenterY.Value = (decimal)((minTop + maxBottom) / 2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"中心座標の取得に失敗しました: {ex.Message}", "エラー");
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            CenterX = (float)numCenterX.Value;
            CenterY = (float)numCenterY.Value;
            Radius = (float)numRadius.Value;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                numCenterX?.Dispose();
                numCenterY?.Dispose();
                numRadius?.Dispose();
                btnUseCurrentCenter?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
