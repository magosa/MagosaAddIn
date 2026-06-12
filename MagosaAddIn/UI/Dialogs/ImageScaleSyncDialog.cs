using MagosaAddIn.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 画像倍率同期ダイアログ。
    /// 2つの画像の実寸法を基準にスケール係数を計算し、画像②をリサイズする。
    /// </summary>
    public class ImageScaleSyncDialog : BaseDialog
    {
        #region フィールド

        private readonly List<PowerPoint.Shape> _imageShapes;
        private readonly ImageScaler _scaler = new ImageScaler();

        // 画像① UI
        private ComboBox _cboImage1;
        private RadioButton _rbImage1ClickMode;
        private RadioButton _rbImage1ManualMode;
        private RadioButton _rbImage1Free;
        private RadioButton _rbImage1Horizontal;
        private RadioButton _rbImage1Vertical;
        private TextBox _txtImage1StartX;
        private TextBox _txtImage1StartY;
        private TextBox _txtImage1EndX;
        private TextBox _txtImage1EndY;
        private Label _lblImage1PixelLength;
        private NumericUpDown _numImage1RealLength;
        private ComboBox _cboImage1Unit;
        private Label _lblImage1Ratio;
        private PictureBox _picImage1Preview;

        // 画像② UI
        private ComboBox _cboImage2;
        private RadioButton _rbImage2ClickMode;
        private RadioButton _rbImage2ManualMode;
        private RadioButton _rbImage2Free;
        private RadioButton _rbImage2Horizontal;
        private RadioButton _rbImage2Vertical;
        private TextBox _txtImage2StartX;
        private TextBox _txtImage2StartY;
        private TextBox _txtImage2EndX;
        private TextBox _txtImage2EndY;
        private Label _lblImage2PixelLength;
        private NumericUpDown _numImage2RealLength;
        private ComboBox _cboImage2Unit;
        private Label _lblImage2Ratio;
        private PictureBox _picImage2Preview;

        // 結果・オプション UI
        private Label _lblScaleFactor;
        private Label _lblCurrentSize;
        private Label _lblNewSize;
        private CheckBox _chkKeepCenter;
        private CheckBox _chkKeepTopLeft;
        private Button _btnCalc;
        private Button _btnExecute;
        private Button _btnCancel;

        // プレビュークリック状態管理
        private bool _image1ClickingStart = true;  // true=次クリックは起点
        private bool _image2ClickingStart = true;

        // 変更通知を一時停止するフラグ
        private bool _suppressUpdate = false;

        // ビットマップ解像度（UpdatePreviewで記録）
        private int _image1BitmapWidth  = 0;
        private int _image2BitmapWidth  = 0;

        // プレビューズーム
        private Panel _panImage1Preview;
        private Panel _panImage2Preview;
        private Label _lblZoom1;
        private Label _lblZoom2;
        private Button _btnZoomIn1;
        private Button _btnZoomOut1;
        private Button _btnZoomIn2;
        private Button _btnZoomOut2;
        private float _zoom1 = 1f;
        private float _zoom2 = 1f;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="imageShapes">スライド内の画像オブジェクト一覧</param>
        public ImageScaleSyncDialog(List<PowerPoint.Shape> imageShapes)
        {
            _imageShapes = imageShapes ?? throw new ArgumentNullException(nameof(imageShapes));
            BuildUI();
            LoadImagesIntoComboBoxes();
        }

        #endregion

        #region UI構築

        private void BuildUI()
        {
            ConfigureForm("画像倍率同期", 1100, 735);

            const int LabelW  = 85;
            const int CtrlX   = 100;
            const int GrpW    = 535;   // 各グループ幅
            const int InnerW  = 515;   // グループ内コントロール幅
            const int PreviewH = 210;  // プレビューパネル高さ
            const int GrpLeft1 = 10;
            const int GrpLeft2 = 555;  // 10 + 535 + 10
            const int GrpTop   = 20;

            // ─── 画像① グループ ─────────────────────────────────
            var grp1 = CreateGroupBox("画像①（基準画像）", new Point(GrpLeft1, GrpTop), new Size(GrpW, 100));
            this.Controls.Add(grp1);

            int gy = 20;
            grp1.Controls.Add(CreateLabel("選択:", new Point(10, gy), LabelW));
            _cboImage1 = new ComboBox { Location = new Point(CtrlX, gy - 2), Width = 395, DropDownStyle = ComboBoxStyle.DropDownList };
            grp1.Controls.Add(_cboImage1);
            gy += 28;

            grp1.Controls.Add(CreateLabel("測定方向:", new Point(10, gy), LabelW));
            _rbImage1Free       = new RadioButton { Text = "自由",    Location = new Point(CtrlX,       gy), AutoSize = true, Checked = true };
            _rbImage1Horizontal = new RadioButton { Text = "水平のみ", Location = new Point(CtrlX + 50,  gy), AutoSize = true };
            _rbImage1Vertical   = new RadioButton { Text = "垂直のみ", Location = new Point(CtrlX + 120, gy), AutoSize = true };
            grp1.Controls.AddRange(new Control[] { _rbImage1Free, _rbImage1Horizontal, _rbImage1Vertical });
            gy += 26;

            grp1.Controls.Add(CreateLabel("入力方式:", new Point(10, gy), LabelW));
            _rbImage1ClickMode  = new RadioButton { Text = "クリック指定", Location = new Point(CtrlX,      gy), AutoSize = true, Checked = true };
            _rbImage1ManualMode = new RadioButton { Text = "数値入力",     Location = new Point(CtrlX + 95, gy), AutoSize = true };
            grp1.Controls.AddRange(new Control[] { _rbImage1ClickMode, _rbImage1ManualMode });
            gy += 26;

            grp1.Controls.Add(CreateLabel("起点 X:", new Point(10, gy), 60));
            _txtImage1StartX = CreateCoordTextBox(new Point(70, gy));
            grp1.Controls.Add(CreateLabel("Y:", new Point(170, gy), 20));
            _txtImage1StartY = CreateCoordTextBox(new Point(190, gy));
            grp1.Controls.AddRange(new Control[] { _txtImage1StartX, _txtImage1StartY });
            gy += 24;

            grp1.Controls.Add(CreateLabel("終点 X:", new Point(10, gy), 60));
            _txtImage1EndX = CreateCoordTextBox(new Point(70, gy));
            grp1.Controls.Add(CreateLabel("Y:", new Point(170, gy), 20));
            _txtImage1EndY = CreateCoordTextBox(new Point(190, gy));
            grp1.Controls.AddRange(new Control[] { _txtImage1EndX, _txtImage1EndY });
            gy += 24;

            _lblImage1PixelLength = new Label { Text = "ピクセル長: 0.00 px", Location = new Point(10, gy), Size = new Size(InnerW, 18), ForeColor = Color.DarkBlue };
            grp1.Controls.Add(_lblImage1PixelLength);
            gy += 22;

            grp1.Controls.Add(CreateLabel("実寸法:", new Point(10, gy), LabelW));
            _numImage1RealLength = CreateNumericUpDown(0.001m, 99999m, 1m, new Point(CtrlX, gy), 3, 1m);
            _numImage1RealLength.Width = 90;
            _cboImage1Unit = new ComboBox { Location = new Point(CtrlX + 95, gy - 2), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboImage1Unit.Items.AddRange(new object[] { "mm", "cm" });
            _cboImage1Unit.SelectedIndex = 1;
            grp1.Controls.AddRange(new Control[] { _numImage1RealLength, _cboImage1Unit });
            gy += 26;

            _lblImage1Ratio = new Label { Text = "倍率: -- mm/px", Location = new Point(10, gy), Size = new Size(InnerW, 18), ForeColor = Color.DarkGreen };
            grp1.Controls.Add(_lblImage1Ratio);
            gy += 22;

            // プレビューヘッダ + ズームボタン
            grp1.Controls.Add(CreateLabel("プレビュー:", new Point(10, gy), LabelW));
            _btnZoomIn1  = new Button { Text = "+",    Location = new Point(InnerW - 90, gy - 2), Size = new Size(28, 22) };
            _btnZoomOut1 = new Button { Text = "−",   Location = new Point(InnerW - 60, gy - 2), Size = new Size(28, 22) };
            _lblZoom1    = new Label  { Text = "100%", Location = new Point(InnerW - 30, gy),     Size = new Size(40, 18), TextAlign = ContentAlignment.MiddleRight };
            grp1.Controls.AddRange(new Control[] { _btnZoomIn1, _btnZoomOut1, _lblZoom1 });
            gy += 24;

            _panImage1Preview = new Panel
            {
                Location = new Point(10, gy),
                Size = new Size(InnerW - 5, PreviewH),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.DimGray
            };
            _picImage1Preview = new PictureBox
            {
                Location = Point.Empty,
                Size = new Size(InnerW - 5, PreviewH),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.LightGray,
                Cursor = Cursors.Cross
            };
            _panImage1Preview.Controls.Add(_picImage1Preview);
            grp1.Controls.Add(_panImage1Preview);
            gy += PreviewH + 4;
            grp1.Height = gy + 10;

            // ─── 画像② グループ ─────────────────────────────────
            var grp2 = CreateGroupBox("画像②（対象画像）", new Point(GrpLeft2, GrpTop), new Size(GrpW, 100));
            this.Controls.Add(grp2);

            gy = 20;
            grp2.Controls.Add(CreateLabel("選択:", new Point(10, gy), LabelW));
            _cboImage2 = new ComboBox { Location = new Point(CtrlX, gy - 2), Width = 395, DropDownStyle = ComboBoxStyle.DropDownList };
            grp2.Controls.Add(_cboImage2);
            gy += 28;

            grp2.Controls.Add(CreateLabel("測定方向:", new Point(10, gy), LabelW));
            _rbImage2Free       = new RadioButton { Text = "自由",    Location = new Point(CtrlX,       gy), AutoSize = true, Checked = true };
            _rbImage2Horizontal = new RadioButton { Text = "水平のみ", Location = new Point(CtrlX + 50,  gy), AutoSize = true };
            _rbImage2Vertical   = new RadioButton { Text = "垂直のみ", Location = new Point(CtrlX + 120, gy), AutoSize = true };
            grp2.Controls.AddRange(new Control[] { _rbImage2Free, _rbImage2Horizontal, _rbImage2Vertical });
            gy += 26;

            grp2.Controls.Add(CreateLabel("入力方式:", new Point(10, gy), LabelW));
            _rbImage2ClickMode  = new RadioButton { Text = "クリック指定", Location = new Point(CtrlX,      gy), AutoSize = true, Checked = true };
            _rbImage2ManualMode = new RadioButton { Text = "数値入力",     Location = new Point(CtrlX + 95, gy), AutoSize = true };
            grp2.Controls.AddRange(new Control[] { _rbImage2ClickMode, _rbImage2ManualMode });
            gy += 26;

            grp2.Controls.Add(CreateLabel("起点 X:", new Point(10, gy), 60));
            _txtImage2StartX = CreateCoordTextBox(new Point(70, gy));
            grp2.Controls.Add(CreateLabel("Y:", new Point(170, gy), 20));
            _txtImage2StartY = CreateCoordTextBox(new Point(190, gy));
            grp2.Controls.AddRange(new Control[] { _txtImage2StartX, _txtImage2StartY });
            gy += 24;

            grp2.Controls.Add(CreateLabel("終点 X:", new Point(10, gy), 60));
            _txtImage2EndX = CreateCoordTextBox(new Point(70, gy));
            grp2.Controls.Add(CreateLabel("Y:", new Point(170, gy), 20));
            _txtImage2EndY = CreateCoordTextBox(new Point(190, gy));
            grp2.Controls.AddRange(new Control[] { _txtImage2EndX, _txtImage2EndY });
            gy += 24;

            _lblImage2PixelLength = new Label { Text = "ピクセル長: 0.00 px", Location = new Point(10, gy), Size = new Size(InnerW, 18), ForeColor = Color.DarkBlue };
            grp2.Controls.Add(_lblImage2PixelLength);
            gy += 22;

            grp2.Controls.Add(CreateLabel("実寸法:", new Point(10, gy), LabelW));
            _numImage2RealLength = CreateNumericUpDown(0.001m, 99999m, 1m, new Point(CtrlX, gy), 3, 1m);
            _numImage2RealLength.Width = 90;
            _cboImage2Unit = new ComboBox { Location = new Point(CtrlX + 95, gy - 2), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboImage2Unit.Items.AddRange(new object[] { "mm", "cm" });
            _cboImage2Unit.SelectedIndex = 1;
            grp2.Controls.AddRange(new Control[] { _numImage2RealLength, _cboImage2Unit });
            gy += 26;

            _lblImage2Ratio = new Label { Text = "倍率: -- mm/px", Location = new Point(10, gy), Size = new Size(InnerW, 18), ForeColor = Color.DarkGreen };
            grp2.Controls.Add(_lblImage2Ratio);
            gy += 22;

            grp2.Controls.Add(CreateLabel("プレビュー:", new Point(10, gy), LabelW));
            _btnZoomIn2  = new Button { Text = "+",    Location = new Point(InnerW - 90, gy - 2), Size = new Size(28, 22) };
            _btnZoomOut2 = new Button { Text = "−",   Location = new Point(InnerW - 60, gy - 2), Size = new Size(28, 22) };
            _lblZoom2    = new Label  { Text = "100%", Location = new Point(InnerW - 30, gy),     Size = new Size(40, 18), TextAlign = ContentAlignment.MiddleRight };
            grp2.Controls.AddRange(new Control[] { _btnZoomIn2, _btnZoomOut2, _lblZoom2 });
            gy += 24;

            _panImage2Preview = new Panel
            {
                Location = new Point(10, gy),
                Size = new Size(InnerW - 5, PreviewH),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.DimGray
            };
            _picImage2Preview = new PictureBox
            {
                Location = Point.Empty,
                Size = new Size(InnerW - 5, PreviewH),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.LightGray,
                Cursor = Cursors.Cross
            };
            _panImage2Preview.Controls.Add(_picImage2Preview);
            grp2.Controls.Add(_panImage2Preview);
            gy += PreviewH + 4;
            grp2.Height = gy + 10;

            // ─── 計算結果 ────────────────────────────────────────
            const int ResultW = 1080;
            int resultY = GrpTop + Math.Max(grp1.Height, grp2.Height) + 8;
            var grpResult = CreateGroupBox("計算結果", new Point(10, resultY), new Size(ResultW, 100));
            this.Controls.Add(grpResult);

            gy = 18;
            _lblScaleFactor = new Label { Text = "スケール係数: --", Location = new Point(10, gy), Size = new Size(ResultW - 20, 18), ForeColor = Color.DarkRed, Font = BoldFont };
            grpResult.Controls.Add(_lblScaleFactor);
            gy += 20;
            _lblCurrentSize = new Label { Text = "現在の画像②サイズ: --", Location = new Point(10, gy), Size = new Size(ResultW - 20, 18) };
            grpResult.Controls.Add(_lblCurrentSize);
            gy += 20;
            _lblNewSize = new Label { Text = "計算後のサイズ: --", Location = new Point(10, gy), Size = new Size(ResultW - 20, 18) };
            grpResult.Controls.Add(_lblNewSize);
            gy += 24;
            _btnCalc = new Button { Text = "再計算", Location = new Point(ResultW - 90, gy - 8), Size = new Size(70, 24) };
            grpResult.Controls.Add(_btnCalc);
            grpResult.Height = gy + 16;

            // ─── オプション ───────────────────────────────────────
            int optY = resultY + grpResult.Height + 8;
            var grpOpt = CreateGroupBox("オプション", new Point(10, optY), new Size(ResultW, 62));
            this.Controls.Add(grpOpt);

            _chkKeepCenter  = new CheckBox { Text = "中心位置を保持する", Location = new Point(10,  18), AutoSize = true, Checked = true };
            _chkKeepTopLeft = new CheckBox { Text = "左上位置を保持する", Location = new Point(200, 18), AutoSize = true };
            grpOpt.Controls.AddRange(new Control[] { _chkKeepCenter, _chkKeepTopLeft });

            // ─── ボタン ────────────────────────────────────────────
            int btnY = optY + grpOpt.Height + 8;
            _btnExecute = new Button { Text = "実行",       Location = new Point(ResultW - 170, btnY), Size = new Size(75, 26) };
            _btnCancel  = new Button { Text = "キャンセル", Location = new Point(ResultW - 88,  btnY), Size = new Size(88, 26), DialogResult = DialogResult.Cancel };
            this.Controls.AddRange(new Control[] { _btnExecute, _btnCancel });

            this.Height = btnY + 66;
            this.CancelButton = _btnCancel;

            WireEvents();
        }

        private TextBox CreateCoordTextBox(Point location)
        {
            return new TextBox { Location = location, Width = 80, Text = "0" };
        }

        private void WireEvents()
        {
            // コンボボックス
            _cboImage1.SelectedIndexChanged += (s, e) => OnImage1SelectionChanged();
            _cboImage2.SelectedIndexChanged += (s, e) => OnImage2SelectionChanged();

            // 測定方向ラジオ
            _rbImage1Free.CheckedChanged += (s, e) => RecalculateImage1();
            _rbImage1Horizontal.CheckedChanged += (s, e) => RecalculateImage1();
            _rbImage1Vertical.CheckedChanged += (s, e) => RecalculateImage1();
            _rbImage2Free.CheckedChanged += (s, e) => RecalculateImage2();
            _rbImage2Horizontal.CheckedChanged += (s, e) => RecalculateImage2();
            _rbImage2Vertical.CheckedChanged += (s, e) => RecalculateImage2();

            // 座標テキスト
            foreach (var txt in new[] { _txtImage1StartX, _txtImage1StartY, _txtImage1EndX, _txtImage1EndY })
                txt.TextChanged += (s, e) => RecalculateImage1();
            foreach (var txt in new[] { _txtImage2StartX, _txtImage2StartY, _txtImage2EndX, _txtImage2EndY })
                txt.TextChanged += (s, e) => RecalculateImage2();

            // 実寸法・単位
            _numImage1RealLength.ValueChanged += (s, e) => RecalculateAll();
            _cboImage1Unit.SelectedIndexChanged += (s, e) => RecalculateAll();
            _numImage2RealLength.ValueChanged += (s, e) => RecalculateAll();
            _cboImage2Unit.SelectedIndexChanged += (s, e) => RecalculateAll();

            // プレビュークリック
            _picImage1Preview.MouseClick += PicImage1Preview_MouseClick;
            _picImage2Preview.MouseClick += PicImage2Preview_MouseClick;
            _picImage1Preview.Paint += PicImage1Preview_Paint;
            _picImage2Preview.Paint += PicImage2Preview_Paint;

            // 入力方式切替
            _rbImage1ClickMode.CheckedChanged += (s, e) => UpdateCoordInputEnabled(1);
            _rbImage1ManualMode.CheckedChanged += (s, e) => UpdateCoordInputEnabled(1);
            _rbImage2ClickMode.CheckedChanged += (s, e) => UpdateCoordInputEnabled(2);
            _rbImage2ManualMode.CheckedChanged += (s, e) => UpdateCoordInputEnabled(2);

            // オプション相互排他
            _chkKeepCenter.CheckedChanged += ChkKeepCenter_CheckedChanged;
            _chkKeepTopLeft.CheckedChanged += ChkKeepTopLeft_CheckedChanged;

            // ボタン
            _btnCalc.Click += (s, e) => RecalculateAll();
            _btnExecute.Click += BtnExecute_Click;

            // ズームボタン
            _btnZoomIn1.Click  += (s, e) => ZoomPreview(1, 0.25f);
            _btnZoomOut1.Click += (s, e) => ZoomPreview(1, -0.25f);
            _btnZoomIn2.Click  += (s, e) => ZoomPreview(2, 0.25f);
            _btnZoomOut2.Click += (s, e) => ZoomPreview(2, -0.25f);

            // マウスホイールズーム（パネル上でスクロール中）
            _panImage1Preview.MouseWheel += (s, e) => ZoomPreview(1, e.Delta > 0 ? 0.25f : -0.25f);
            _panImage2Preview.MouseWheel += (s, e) => ZoomPreview(2, e.Delta > 0 ? 0.25f : -0.25f);
            _picImage1Preview.MouseWheel += (s, e) => ZoomPreview(1, e.Delta > 0 ? 0.25f : -0.25f);
            _picImage2Preview.MouseWheel += (s, e) => ZoomPreview(2, e.Delta > 0 ? 0.25f : -0.25f);
        }

        #endregion

        #region 初期化

        private void LoadImagesIntoComboBoxes()
        {
            _suppressUpdate = true;
            try
            {
                _cboImage1.Items.Clear();
                _cboImage2.Items.Clear();
                foreach (var shape in _imageShapes)
                {
                    _cboImage1.Items.Add(shape.Name);
                    _cboImage2.Items.Add(shape.Name);
                }
                if (_cboImage1.Items.Count > 0) _cboImage1.SelectedIndex = 0;
                if (_cboImage2.Items.Count > 1) _cboImage2.SelectedIndex = 1;
                else if (_cboImage2.Items.Count > 0) _cboImage2.SelectedIndex = 0;
            }
            finally
            {
                _suppressUpdate = false;
            }

            UpdateCoordInputEnabled(1);
            UpdateCoordInputEnabled(2);
            OnImage1SelectionChanged();
            OnImage2SelectionChanged();
        }

        #endregion

        #region 画像選択変更

        private void OnImage1SelectionChanged()
        {
            _image1ClickingStart = true;
            UpdatePreview(1);
            RecalculateImage1();
        }

        private void OnImage2SelectionChanged()
        {
            _image2ClickingStart = true;
            UpdatePreview(2);
            RecalculateImage2();
            UpdateCurrentSizeLabel();
        }

        #endregion

        #region プレビュー更新

        private void UpdatePreview(int imageIndex)
        {
            var pic = imageIndex == 1 ? _picImage1Preview : _picImage2Preview;
            var shape = GetSelectedShape(imageIndex);
            if (shape == null)
            {
                pic.Image = null;
                if (imageIndex == 1) _zoom1 = 1f; else _zoom2 = 1f;
                UpdatePictureBoxSize(imageIndex);
                return;
            }

            try
            {
                var bmp = ExtractShapeImage(shape);
                pic.Image = bmp;
                if (imageIndex == 1)
                {
                    _zoom1 = 1f;
                    _image1BitmapWidth = bmp?.Width ?? 0;
                }
                else
                {
                    _zoom2 = 1f;
                    _image2BitmapWidth = bmp?.Width ?? 0;
                }
                UpdatePictureBoxSize(imageIndex);
            }
            catch
            {
                pic.Image = null;
                if (imageIndex == 1) { _zoom1 = 1f; _image1BitmapWidth = 0; }
                else { _zoom2 = 1f; _image2BitmapWidth = 0; }
                UpdatePictureBoxSize(imageIndex);
            }
        }

        /// <summary>
        /// PowerPoint Shape から Bitmap を生成する（クリップボード経由）
        /// </summary>
        private Bitmap ExtractShapeImage(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(() =>
            {
                shape.Copy();
                if (Clipboard.ContainsImage())
                {
                    var img = Clipboard.GetImage();
                    if (img != null)
                        return new Bitmap(img);
                }
                return null;
            }, "画像プレビュー取得", defaultValue: null, suppressErrors: true);
        }

        private void PicImage1Preview_Paint(object sender, PaintEventArgs e)
        {
            DrawMeasurementLine(e.Graphics, _picImage1Preview, 1);
        }

        private void PicImage2Preview_Paint(object sender, PaintEventArgs e)
        {
            DrawMeasurementLine(e.Graphics, _picImage2Preview, 2);
        }

        private void DrawMeasurementLine(Graphics g, PictureBox pic, int imageIndex)
        {
            if (!TryParseCoords(imageIndex, out float x1, out float y1, out float x2, out float y2))
                return;
            if (pic.Image == null)
                return;

            var imgSize = pic.Image.Size;
            var picSize = pic.ClientSize;

            // ZoomモードでのImageRect計算
            float scaleX = (float)picSize.Width / imgSize.Width;
            float scaleY = (float)picSize.Height / imgSize.Height;
            float scale = Math.Min(scaleX, scaleY);
            float offsetX = (picSize.Width - imgSize.Width * scale) / 2f;
            float offsetY = (picSize.Height - imgSize.Height * scale) / 2f;

            // ピクセル座標 → 表示座標
            float px1 = x1 * scale + offsetX;
            float py1 = y1 * scale + offsetY;
            float px2 = x2 * scale + offsetX;
            float py2 = y2 * scale + offsetY;

            using (var pen = new Pen(Color.Red, 2f))
            {
                g.DrawLine(pen, px1, py1, px2, py2);
                g.FillEllipse(Brushes.LimeGreen, px1 - 4, py1 - 4, 8, 8);
                g.FillEllipse(Brushes.Red, px2 - 4, py2 - 4, 8, 8);
            }
        }

        #endregion

        #region プレビューズーム

        /// <summary>
        /// プレビューのズームレベルを変更する
        /// </summary>
        private void ZoomPreview(int imageIndex, float delta)
        {
            float zoom = imageIndex == 1 ? _zoom1 : _zoom2;
            zoom = Math.Max(0.25f, Math.Min(8f, zoom + delta));
            if (imageIndex == 1) _zoom1 = zoom;
            else _zoom2 = zoom;
            UpdatePictureBoxSize(imageIndex);
        }

        /// <summary>
        /// ズームに合わせてPictureBoxのサイズを更新する。
        /// PictureBox(SizeMode=Zoom)はパネルのフィットサイズを基準に拡縮し、
        /// パネルのAutoScrollでスクロール可能になる。
        /// </summary>
        private void UpdatePictureBoxSize(int imageIndex)
        {
            var pic  = imageIndex == 1 ? _picImage1Preview : _picImage2Preview;
            var pan  = imageIndex == 1 ? _panImage1Preview : _panImage2Preview;
            var lbl  = imageIndex == 1 ? _lblZoom1 : _lblZoom2;
            float zoom = imageIndex == 1 ? _zoom1 : _zoom2;

            lbl.Text = $"{(int)(zoom * 100)}%";

            if (pic.Image == null)
            {
                pic.Size = pan.ClientSize;
                return;
            }

            var imgSize  = pic.Image.Size;
            var panelSize = pan.ClientSize;

            // zoom=1.0 のときにパネルにフィットするスケールを計算
            float fitScaleX = (float)panelSize.Width  / imgSize.Width;
            float fitScaleY = (float)panelSize.Height / imgSize.Height;
            float fitScale  = Math.Min(fitScaleX, fitScaleY);

            int newW = Math.Max(1, (int)(imgSize.Width  * fitScale * zoom));
            int newH = Math.Max(1, (int)(imgSize.Height * fitScale * zoom));
            pic.Size = new Size(newW, newH);
            pic.Invalidate();
        }

        #endregion

        #region プレビュークリック座標取得

        private void PicImage1Preview_MouseClick(object sender, MouseEventArgs e)
        {
            if (!_rbImage1ClickMode.Checked) return;
            var (px, py) = ConvertPreviewClickToPixel(_picImage1Preview, e.X, e.Y);
            if (_image1ClickingStart)
            {
                SetTextSafe(_txtImage1StartX, px.ToString("F0"));
                SetTextSafe(_txtImage1StartY, py.ToString("F0"));
            }
            else
            {
                SetTextSafe(_txtImage1EndX, px.ToString("F0"));
                SetTextSafe(_txtImage1EndY, py.ToString("F0"));
            }
            _image1ClickingStart = !_image1ClickingStart;
            _picImage1Preview.Invalidate();
        }

        private void PicImage2Preview_MouseClick(object sender, MouseEventArgs e)
        {
            if (!_rbImage2ClickMode.Checked) return;
            var (px, py) = ConvertPreviewClickToPixel(_picImage2Preview, e.X, e.Y);
            if (_image2ClickingStart)
            {
                SetTextSafe(_txtImage2StartX, px.ToString("F0"));
                SetTextSafe(_txtImage2StartY, py.ToString("F0"));
            }
            else
            {
                SetTextSafe(_txtImage2EndX, px.ToString("F0"));
                SetTextSafe(_txtImage2EndY, py.ToString("F0"));
            }
            _image2ClickingStart = !_image2ClickingStart;
            _picImage2Preview.Invalidate();
        }

        private (float X, float Y) ConvertPreviewClickToPixel(PictureBox pic, int clickX, int clickY)
        {
            if (pic.Image == null) return (0f, 0f);

            var imgSize = pic.Image.Size;
            var picSize = pic.ClientSize;
            float scaleX = (float)picSize.Width / imgSize.Width;
            float scaleY = (float)picSize.Height / imgSize.Height;
            float scale = Math.Min(scaleX, scaleY);
            float offsetX = (picSize.Width - imgSize.Width * scale) / 2f;
            float offsetY = (picSize.Height - imgSize.Height * scale) / 2f;

            float px = (clickX - offsetX) / scale;
            float py = (clickY - offsetY) / scale;

            px = Math.Max(0f, Math.Min(px, imgSize.Width));
            py = Math.Max(0f, Math.Min(py, imgSize.Height));
            return (px, py);
        }

        private void SetTextSafe(TextBox txt, string value)
        {
            _suppressUpdate = true;
            txt.Text = value;
            _suppressUpdate = false;
        }

        #endregion

        #region 入力方式切替

        private void UpdateCoordInputEnabled(int imageIndex)
        {
            bool isManual = imageIndex == 1
                ? _rbImage1ManualMode.Checked
                : _rbImage2ManualMode.Checked;

            var coordBoxes = imageIndex == 1
                ? new[] { _txtImage1StartX, _txtImage1StartY, _txtImage1EndX, _txtImage1EndY }
                : new[] { _txtImage2StartX, _txtImage2StartY, _txtImage2EndX, _txtImage2EndY };

            foreach (var txt in coordBoxes)
                txt.ReadOnly = !isManual;
        }

        #endregion

        #region リアルタイム計算

        private void RecalculateImage1()
        {
            if (_suppressUpdate) return;
            if (!TryParseCoords(1, out float x1, out float y1, out float x2, out float y2))
            {
                _lblImage1PixelLength.Text = "ピクセル長: -- px";
                _lblImage1Ratio.Text = "倍率: --";
                RecalculateScaleFactor();
                return;
            }

            var mode = GetMeasurementMode(1);
            float length = _scaler.CalculateDistance(x1, y1, x2, y2, mode);
            _lblImage1PixelLength.Text = $"ピクセル長: {length:F2} px";

            UpdateRatioLabel(1, length);
            RecalculateScaleFactor();
            _picImage1Preview.Invalidate();
        }

        private void RecalculateImage2()
        {
            if (_suppressUpdate) return;
            if (!TryParseCoords(2, out float x1, out float y1, out float x2, out float y2))
            {
                _lblImage2PixelLength.Text = "ピクセル長: -- px";
                _lblImage2Ratio.Text = "倍率: --";
                RecalculateScaleFactor();
                return;
            }

            var mode = GetMeasurementMode(2);
            float length = _scaler.CalculateDistance(x1, y1, x2, y2, mode);
            _lblImage2PixelLength.Text = $"ピクセル長: {length:F2} px";

            UpdateRatioLabel(2, length);
            RecalculateScaleFactor();
            _picImage2Preview.Invalidate();
        }

        private void RecalculateAll()
        {
            if (_suppressUpdate) return;
            RecalculateImage1();
            RecalculateImage2();
        }

        private void UpdateRatioLabel(int imageIndex, float pixelLength)
        {
            var num = imageIndex == 1 ? _numImage1RealLength : _numImage2RealLength;
            var cboUnit = imageIndex == 1 ? _cboImage1Unit : _cboImage2Unit;
            var lbl = imageIndex == 1 ? _lblImage1Ratio : _lblImage2Ratio;

            float realLength = (float)num.Value;
            var unit = GetSizeUnit(cboUnit);
            string unitStr = cboUnit.SelectedItem?.ToString() ?? "cm";

            if (realLength <= 0f || pixelLength < 0.001f)
            {
                lbl.Text = "倍率: --";
                return;
            }

            try
            {
                float ratio = _scaler.CalculateImageRatio(realLength, unit, pixelLength);
                float ratioMm = _scaler.NormalizeRealLengthToMillimeter(realLength, unit);
                lbl.Text = $"倍率: {realLength} {unitStr} / {pixelLength:F2} px = {ratioMm / pixelLength:F4} mm/px";
            }
            catch
            {
                lbl.Text = "倍率: --";
            }
        }

        private void RecalculateScaleFactor()
        {
            _lblScaleFactor.Text = "スケール係数: --";
            _lblNewSize.Text = "計算後のサイズ: --";

            try
            {
                if (!TryParseCoords(1, out float x1s, out float y1s, out float x1e, out float y1e)) return;
                if (!TryParseCoords(2, out float x2s, out float y2s, out float x2e, out float y2e)) return;

                float r1 = (float)_numImage1RealLength.Value;
                float r2 = (float)_numImage2RealLength.Value;
                if (r1 <= 0f || r2 <= 0f) return;

                var shape2 = GetSelectedShape(2);
                if (shape2 == null) return;

                var unit1 = GetSizeUnit(_cboImage1Unit);
                var unit2 = GetSizeUnit(_cboImage2Unit);
                var mode1 = GetMeasurementMode(1);
                var mode2 = GetMeasurementMode(2);

                float l1 = _scaler.CalculateDistance(x1s, y1s, x1e, y1e, mode1);
                float l2 = _scaler.CalculateDistance(x2s, y2s, x2e, y2e, mode2);
                if (l1 < 0.001f || l2 < 0.001f) return;

                float ratio1 = _scaler.CalculateImageRatio(r1, unit1, l1);
                float ratio2 = _scaler.CalculateImageRatio(r2, unit2, l2);
                if (Math.Abs(ratio1) < 1e-9f) return;

                // ビットマップ解像度補正込みスケール係数
                float scale = ratio2 / ratio1;
                if (_image1BitmapWidth > 0 && _image2BitmapWidth > 0)
                {
                    var sh1 = GetSelectedShape(1);
                    var sh2 = GetSelectedShape(2);
                    if (sh1 != null && sh2 != null && sh1.Width > 0.001f && sh2.Width > 0.001f)
                        scale *= ((float)_image2BitmapWidth * sh1.Width)
                               / ((float)_image1BitmapWidth * sh2.Width);
                }
                _lblScaleFactor.Text = $"スケール係数: {scale:F4} 倍";

                float newW = shape2.Width * scale;
                float newH = shape2.Height * scale;
                _lblNewSize.Text = $"計算後のサイズ: {newW:F1} × {newH:F1} pt";
            }
            catch { /* 入力途中の例外は無視 */ }
        }

        private void UpdateCurrentSizeLabel()
        {
            var shape2 = GetSelectedShape(2);
            if (shape2 == null)
            {
                _lblCurrentSize.Text = "現在の画像②サイズ: --";
                return;
            }
            try
            {
                _lblCurrentSize.Text = $"現在の画像②サイズ: {shape2.Width:F1} × {shape2.Height:F1} pt";
            }
            catch
            {
                _lblCurrentSize.Text = "現在の画像②サイズ: --";
            }
        }

        #endregion

        #region オプション相互排他

        private void ChkKeepCenter_CheckedChanged(object sender, EventArgs e)
        {
            if (_chkKeepCenter.Checked && _chkKeepTopLeft.Checked)
                _chkKeepTopLeft.Checked = false;
        }

        private void ChkKeepTopLeft_CheckedChanged(object sender, EventArgs e)
        {
            if (_chkKeepTopLeft.Checked && _chkKeepCenter.Checked)
                _chkKeepCenter.Checked = false;
        }

        #endregion

        #region 実行

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            var shape1 = GetSelectedShape(1);
            var shape2 = GetSelectedShape(2);

            if (!TryParseCoords(1, out float x1s, out float y1s, out float x1e, out float y1e)) { x1s = 0; y1s = 0; x1e = 0; y1e = 0; }
            if (!TryParseCoords(2, out float x2s, out float y2s, out float x2e, out float y2e)) { x2s = 0; y2s = 0; x2e = 0; y2e = 0; }

            float r1 = (float)_numImage1RealLength.Value;
            float r2 = (float)_numImage2RealLength.Value;
            var mode = GetMeasurementMode(1); // 画像①の測定方向をバリデ―ションに使用

            var (isValid, errorMsg) = _scaler.ValidateInputs(
                shape1, x1s, y1s, x1e, y1e, r1,
                shape2, x2s, y2s, x2e, y2e, r2,
                mode);

            if (!isValid)
            {
                MessageBox.Show(errorMsg, "Magosa Tools - 入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            float scaleFactor;
            try
            {
                scaleFactor = _scaler.CalculateScaleFactor(
                    shape1, x1s, y1s, x1e, y1e, r1, GetSizeUnit(_cboImage1Unit),
                    shape2, x2s, y2s, x2e, y2e, r2, GetSizeUnit(_cboImage2Unit),
                    GetMeasurementMode(1), GetMeasurementMode(2),
                    _image1BitmapWidth, _image2BitmapWidth);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Magosa Tools - 計算エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            float newW = shape2.Width * scaleFactor;
            float newH = shape2.Height * scaleFactor;
            var posMode = _chkKeepCenter.Checked ? ResizeMode.KeepCenter : ResizeMode.KeepTopLeft;
            string posModeStr = posMode == ResizeMode.KeepCenter ? "中心を保持" : "左上を保持";

            string confirmMsg =
                $"スケーリングを実行しますか？\n\n" +
                $"スケール係数: {scaleFactor:F4} 倍\n" +
                $"現在サイズ: {shape2.Width:F1} × {shape2.Height:F1} pt\n" +
                $"計算後サイズ: {newW:F1} × {newH:F1} pt\n" +
                $"位置保持: {posModeStr}";

            var result = MessageBox.Show(confirmMsg, "Magosa Tools - 画像倍率同期",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result != DialogResult.OK) return;

            try
            {
                _scaler.ApplyScaling(shape2, scaleFactor, posMode);
                MessageBox.Show(
                    $"スケーリングが完了しました。\n画像②を {scaleFactor:F4} 倍に拡大縮小しました。",
                    "Magosa Tools - 画像倍率同期",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Magosa Tools - スケーリングエラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region ヘルパー

        private PowerPoint.Shape GetSelectedShape(int imageIndex)
        {
            var cbo = imageIndex == 1 ? _cboImage1 : _cboImage2;
            if (cbo.SelectedIndex < 0 || cbo.SelectedIndex >= _imageShapes.Count)
                return null;
            return _imageShapes[cbo.SelectedIndex];
        }

        private bool TryParseCoords(int imageIndex, out float x1, out float y1, out float x2, out float y2)
        {
            var txts = imageIndex == 1
                ? new[] { _txtImage1StartX, _txtImage1StartY, _txtImage1EndX, _txtImage1EndY }
                : new[] { _txtImage2StartX, _txtImage2StartY, _txtImage2EndX, _txtImage2EndY };

            // 短絡評価を避けて全 out パラメータを必ず代入する
            bool ok1 = float.TryParse(txts[0].Text, out x1);
            bool ok2 = float.TryParse(txts[1].Text, out y1);
            bool ok3 = float.TryParse(txts[2].Text, out x2);
            bool ok4 = float.TryParse(txts[3].Text, out y2);
            return ok1 && ok2 && ok3 && ok4;
        }

        private MeasurementMode GetMeasurementMode(int imageIndex)
        {
            if (imageIndex == 1)
            {
                if (_rbImage1Horizontal.Checked) return MeasurementMode.HorizontalOnly;
                if (_rbImage1Vertical.Checked) return MeasurementMode.VerticalOnly;
                return MeasurementMode.Free;
            }
            else
            {
                if (_rbImage2Horizontal.Checked) return MeasurementMode.HorizontalOnly;
                if (_rbImage2Vertical.Checked) return MeasurementMode.VerticalOnly;
                return MeasurementMode.Free;
            }
        }

        private SizeUnit GetSizeUnit(ComboBox cbo)
        {
            return cbo.SelectedItem?.ToString() == "mm" ? SizeUnit.Millimeter : SizeUnit.Centimeter;
        }

        #endregion
    }
}
