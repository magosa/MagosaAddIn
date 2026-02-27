using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MagosaAddIn.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 図形スタイルライブラリダイアログ
    /// スタイルの一覧表示・保存・適用・削除・インポート/エクスポートを行う
    /// </summary>
    public class StyleLibraryDialog : BaseDialog
    {
        #region フィールド

        private readonly ShapeStyleLibrary _library;
        private readonly List<PowerPoint.Shape> _selectedShapes;

        // フィルター
        private CheckBox _chkFavoriteOnly;
        private TextBox _txtSearch;
        private Label _lblCount;

        // スタイル一覧
        private ListView _lvStyles;

        // プレビューパネル
        private Panel _pnlPreview;
        private Label _lblPreviewFill;
        private Label _lblPreviewLine;
        private Label _lblPreviewShadow;
        private Label _lblPreviewFont;
        private Label _lblPreviewName;
        private Label _lblPreviewDate;

        // ボタン群
        private Button _btnApply;
        private Button _btnSave;
        private Button _btnDelete;
        private Button _btnFavorite;
        private Button _btnExport;
        private Button _btnImport;
        private Button _btnClose;

        // 現在選択中のスタイルエントリ
        private StyleEntry _selectedEntry;

        #endregion

        #region コンストラクタ

        public StyleLibraryDialog(ShapeStyleLibrary library, List<PowerPoint.Shape> selectedShapes)
        {
            _library = library ?? throw new ArgumentNullException(nameof(library));
            _selectedShapes = selectedShapes ?? new List<PowerPoint.Shape>();
            InitializeDialog();
        }

        #endregion

        #region 初期化

        private void InitializeDialog()
        {
            ConfigureForm("図形スタイルライブラリ", 720, 550);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(720, 550);
            BuildUI();
            RefreshList();
        }

        private void BuildUI()
        {
            // ─── 左ペイン：フィルター＋一覧 ─────────────────────
            var lblFilter = new Label
            {
                Text = "検索:",
                Location = new Point(10, 14),
                Size = new Size(40, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _txtSearch = new TextBox
            {
                Location = new Point(52, 12),
                Size = new Size(180, 22)
            };
            _txtSearch.TextChanged += (s, e) => RefreshList();

            _chkFavoriteOnly = new CheckBox
            {
                Text = "お気に入りのみ",
                Location = new Point(244, 13),
                Size = new Size(135, 20),
                Checked = false
            };
            _chkFavoriteOnly.CheckedChanged += (s, e) => RefreshList();

            _lblCount = new Label
            {
                Text = "0件",
                Location = new Point(390, 14),
                Size = new Size(60, 20),
                ForeColor = Color.Gray
            };

            // ListView
            _lvStyles = new ListView
            {
                Location = new Point(10, 40),
                Size = new Size(380, 380),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false
            };
            _lvStyles.Columns.Add("", 26);        // お気に入り
            _lvStyles.Columns.Add("スタイル名", 150);
            _lvStyles.Columns.Add("塗り", 50);
            _lvStyles.Columns.Add("枠", 30);
            _lvStyles.Columns.Add("影", 30);
            _lvStyles.Columns.Add("登録日", 110);
            _lvStyles.SelectedIndexChanged += LvStyles_SelectedIndexChanged;
            _lvStyles.DoubleClick += (s, e) => BtnApply_Click(s, e);

            // ─── 右ペイン：プレビュー ──────────────────────────
            var grpPreview = new GroupBox
            {
                Text = "プレビュー",
                Location = new Point(400, 40),
                Size = new Size(280, 220)
            };

            _pnlPreview = new Panel
            {
                Location = new Point(10, 20),
                Size = new Size(260, 120),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            _pnlPreview.Paint += PnlPreview_Paint;

            _lblPreviewName = new Label
            {
                Location = new Point(10, 148),
                Size = new Size(260, 18),
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                ForeColor = Color.Black
            };
            _lblPreviewFill = new Label { Location = new Point(10, 166), Size = new Size(260, 16), ForeColor = Color.Gray };
            _lblPreviewLine = new Label { Location = new Point(10, 182), Size = new Size(260, 16), ForeColor = Color.Gray };
            _lblPreviewShadow = new Label { Location = new Point(10, 198), Size = new Size(130, 16), ForeColor = Color.Gray };
            _lblPreviewFont = new Label { Location = new Point(140, 198), Size = new Size(130, 16), ForeColor = Color.Gray };
            _lblPreviewDate = new Label { Location = new Point(10, 214), Size = new Size(260, 16), ForeColor = Color.LightGray };

            grpPreview.Controls.AddRange(new Control[]
            {
                _pnlPreview,
                _lblPreviewName, _lblPreviewFill, _lblPreviewLine,
                _lblPreviewShadow, _lblPreviewFont, _lblPreviewDate
            });

            // ─── 右ペイン：ボタン群 ───────────────────────────
            int btnX = 400;
            int btnY = 270;
            int btnW = 130;
            int btnH = 26;
            int btnGap = 4;

            _btnApply = new Button
            {
                Text = "選択図形に適用",
                Location = new Point(btnX, btnY),
                Size = new Size(btnW * 2 + 10, btnH)
            };
            _btnApply.Click += BtnApply_Click;
            btnY += btnH + btnGap + 6;

            var sep1 = new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(btnX, btnY), Size = new Size(btnW * 2 + 10, 2) };
            btnY += 10;

            _btnSave = new Button
            {
                Text = "現在の図形を登録",
                Location = new Point(btnX, btnY),
                Size = new Size(btnW * 2 + 10, btnH)
            };
            _btnSave.Click += BtnSave_Click;
            btnY += btnH + btnGap;

            _btnFavorite = new Button
            {
                Text = "お気に入り切替",
                Location = new Point(btnX, btnY),
                Size = new Size(btnW * 2 + 10, btnH)
            };
            _btnFavorite.Click += BtnFavorite_Click;
            btnY += btnH + btnGap;

            _btnDelete = new Button
            {
                Text = "削除",
                Location = new Point(btnX, btnY),
                Size = new Size(btnW * 2 + 10, btnH),
                ForeColor = Color.FromArgb(180, 0, 0)
            };
            _btnDelete.Click += BtnDelete_Click;
            btnY += btnH + btnGap + 6;

            var sep2 = new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(btnX, btnY), Size = new Size(btnW * 2 + 10, 2) };
            btnY += 10;

            _btnExport = new Button
            {
                Text = "エクスポート...",
                Location = new Point(btnX, btnY),
                Size = new Size(btnW, btnH)
            };
            _btnExport.Click += BtnExport_Click;

            _btnImport = new Button
            {
                Text = "インポート...",
                Location = new Point(btnX + btnW + 10, btnY),
                Size = new Size(btnW, btnH)
            };
            _btnImport.Click += BtnImport_Click;
            btnY += btnH + btnGap + 6;

            // 閉じるボタン
            _btnClose = new Button
            {
                Text = "閉じる",
                Location = new Point(btnX + btnW + 10, btnY),
                Size = new Size(btnW, btnH)
            };
            _btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[]
            {
                lblFilter, _txtSearch, _chkFavoriteOnly, _lblCount,
                _lvStyles, grpPreview,
                _btnApply, sep1,
                _btnSave, _btnFavorite, _btnDelete, sep2,
                _btnExport, _btnImport,
                _btnClose
            });

            // ウィンドウリサイズ対応
            this.Resize += (s, e) => UpdateLayout();
        }

        private void UpdateLayout()
        {
            int w = this.ClientSize.Width;
            int h = this.ClientSize.Height;
            _lvStyles.Height = h - 80;
            _txtSearch.Width = Math.Max(100, w - 340);
        }

        #endregion

        #region スタイル一覧

        private void RefreshList()
        {
            _lvStyles.BeginUpdate();
            _lvStyles.Items.Clear();

            var styles = _chkFavoriteOnly.Checked
                ? _library.GetFavoriteStyles()
                : _library.GetAllStyles();

            // 検索フィルター
            string keyword = _txtSearch.Text.Trim();
            if (!string.IsNullOrEmpty(keyword))
            {
                styles = styles.Where(s =>
                    s.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            }

            // お気に入り→通常の順
            styles = styles
                .OrderByDescending(s => s.IsFavorite)
                .ThenBy(s => s.Name)
                .ToList();

            foreach (var entry in styles)
            {
                var item = new ListViewItem(entry.IsFavorite ? "★" : "");
                item.SubItems.Add(entry.Name);
                item.SubItems.Add(entry.HasFill ? (entry.HasGradient ? "G" : "■") : "-");
                item.SubItems.Add(entry.HasLine ? "─" : "-");
                item.SubItems.Add(entry.HasShadow ? "●" : "-");
                item.SubItems.Add(entry.CreatedAt ?? "");
                item.Tag = entry;

                if (entry.IsFavorite)
                {
                    item.ForeColor = Color.FromArgb(160, 100, 0);
                    item.Font = new Font(_lvStyles.Font, FontStyle.Bold);
                }

                _lvStyles.Items.Add(item);
            }

            _lblCount.Text = $"{styles.Count}件";
            _lvStyles.EndUpdate();

            UpdateButtons();
        }

        private void LvStyles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lvStyles.SelectedItems.Count > 0)
            {
                _selectedEntry = _lvStyles.SelectedItems[0].Tag as StyleEntry;
            }
            else
            {
                _selectedEntry = null;
            }

            UpdatePreview();
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            bool hasSelection = _selectedEntry != null;
            bool hasShapes = _selectedShapes.Count > 0;

            _btnApply.Enabled = hasSelection && hasShapes;
            _btnApply.Text = hasShapes
                ? $"選択図形に適用 ({_selectedShapes.Count}個)"
                : "選択図形に適用";
            _btnFavorite.Enabled = hasSelection;
            _btnDelete.Enabled = hasSelection;
            _btnExport.Enabled = _library.Count > 0;
            _btnSave.Enabled = _selectedShapes.Count == 1;
            _btnSave.Text = _selectedShapes.Count == 1
                ? "現在の図形を登録"
                : "登録（図形を1個選択）";
        }

        #endregion

        #region プレビュー描画

        private void UpdatePreview()
        {
            if (_selectedEntry == null)
            {
                _lblPreviewName.Text = "（スタイルを選択してください）";
                _lblPreviewFill.Text = "";
                _lblPreviewLine.Text = "";
                _lblPreviewShadow.Text = "";
                _lblPreviewFont.Text = "";
                _lblPreviewDate.Text = "";
            }
            else
            {
                var e = _selectedEntry;
                _lblPreviewName.Text = (e.IsFavorite ? "★ " : "") + e.Name;
                _lblPreviewFill.Text = e.HasFill
                    ? (e.HasGradient ? "塗り: グラデーション" : $"塗り: #{e.FillColor:X6}  透明度:{e.FillTransparency:P0}")
                    : "塗り: なし";
                _lblPreviewLine.Text = e.HasLine
                    ? $"枠線: #{e.LineColor:X6}  {e.LineWeight:F1}pt"
                    : "枠線: なし";
                _lblPreviewShadow.Text = e.HasShadow ? "影: あり" : "影: なし";
                _lblPreviewFont.Text = !string.IsNullOrEmpty(e.FontName)
                    ? $"フォント: {e.FontName} {e.FontSize:F0}pt"
                    : "";
                _lblPreviewDate.Text = e.CreatedAt ?? "";
            }

            _pnlPreview.Invalidate();
        }

        private void PnlPreview_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(20, 10, _pnlPreview.Width - 40, _pnlPreview.Height - 20);

            if (_selectedEntry == null)
            {
                g.FillRectangle(Brushes.WhiteSmoke, rect);
                g.DrawRectangle(Pens.LightGray, rect);
                return;
            }

            var entry = _selectedEntry;

            // 影（簡易）
            if (entry.HasShadow)
            {
                var shadowRect = new Rectangle(rect.X + 4, rect.Y + 4, rect.Width, rect.Height);
                var shadowColor = entry.GetLineDrawingColor();
                using (var shadowBrush = new SolidBrush(Color.FromArgb(80, 100, 100, 100)))
                    g.FillRectangle(shadowBrush, shadowRect);
            }

            // 塗りつぶし
            if (entry.HasFill)
            {
                if (entry.HasGradient)
                {
                    var c1 = entry.GetFillDrawingColor();
                    var c2 = Color.FromArgb(
                        entry.GradientColor2 & 0xFF,
                        (entry.GradientColor2 >> 8) & 0xFF,
                        (entry.GradientColor2 >> 16) & 0xFF);
                    using (var brush = new LinearGradientBrush(rect, c1, c2, LinearGradientMode.Horizontal))
                        g.FillRectangle(brush, rect);
                }
                else
                {
                    var fillColor = entry.GetFillDrawingColor();
                    int alpha = (int)((1f - entry.FillTransparency) * 255);
                    alpha = Math.Max(0, Math.Min(255, alpha));
                    using (var brush = new SolidBrush(Color.FromArgb(alpha, fillColor)))
                        g.FillRectangle(brush, rect);
                }
            }
            else
            {
                g.FillRectangle(Brushes.White, rect);
            }

            // 枠線
            if (entry.HasLine)
            {
                var lineColor = entry.GetLineDrawingColor();
                float lw = Math.Max(1f, Math.Min(entry.LineWeight, 8f));
                using (var pen = new Pen(lineColor, lw))
                    g.DrawRectangle(pen, rect);
            }
            else
            {
                g.DrawRectangle(Pens.LightGray, rect);
            }

            // テキストサンプル
            string sampleText = string.IsNullOrEmpty(entry.FontName) ? "Aa" : "Aa";
            var fontColor = entry.FontColor > 0
                ? Color.FromArgb(entry.FontColor & 0xFF, (entry.FontColor >> 8) & 0xFF, (entry.FontColor >> 16) & 0xFF)
                : Color.Black;
            float fontSize = entry.FontSize > 0 ? Math.Min(entry.FontSize, 28f) : 18f;

            try
            {
                string fontFamilyName = !string.IsNullOrEmpty(entry.FontName) ? entry.FontName : "Arial";
                var fontStyle = FontStyle.Regular;
                if (entry.FontBold) fontStyle |= FontStyle.Bold;
                if (entry.FontItalic) fontStyle |= FontStyle.Italic;

                using (var font = new Font(fontFamilyName, fontSize, fontStyle, GraphicsUnit.Point))
                using (var brush = new SolidBrush(fontColor))
                {
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    g.DrawString(sampleText, font, brush, rect, sf);
                }
            }
            catch
            {
                using (var font = new Font("Arial", 18, FontStyle.Regular, GraphicsUnit.Point))
                using (var brush = new SolidBrush(fontColor))
                {
                    var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    g.DrawString("Aa", font, brush, rect, sf);
                }
            }
        }

        #endregion

        #region ボタンハンドラ

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (_selectedEntry == null || _selectedShapes.Count == 0) return;

            try
            {
                int count = _library.ApplyStyleToShapes(_selectedShapes, _selectedEntry.Name);
                ErrorHandler.ShowOperationSuccess("スタイル適用",
                    $"{count}個の図形にスタイル「{_selectedEntry.Name}」を適用しました");
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("スタイル適用", ex);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_selectedShapes.Count != 1)
            {
                MessageBox.Show("スタイルを保存するには、図形を1つだけ選択してください。",
                    "スタイル保存", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var nameDialog = new StyleNameInputDialog(_library))
            {
                if (nameDialog.ShowDialog() != DialogResult.OK) return;
                string name = nameDialog.StyleName;

                try
                {
                    _library.SaveStyleFromShape(_selectedShapes[0], name);
                    ErrorHandler.ShowOperationSuccess("スタイル保存", $"スタイル「{name}」を保存しました");
                    RefreshList();
                }
                catch (Exception ex)
                {
                    ErrorHandler.ShowOperationError("スタイル保存", ex);
                }
            }
        }

        private void BtnFavorite_Click(object sender, EventArgs e)
        {
            if (_selectedEntry == null) return;

            _library.ToggleFavorite(_selectedEntry.Name);
            bool nowFavorite = !_selectedEntry.IsFavorite; // 切り替え後
            string msg = nowFavorite ? "お気に入りに追加しました" : "お気に入りを解除しました";
            RefreshList();

            // 同じスタイルを再選択
            foreach (ListViewItem item in _lvStyles.Items)
            {
                if ((item.Tag as StyleEntry)?.Name == _selectedEntry.Name)
                {
                    item.Selected = true;
                    item.EnsureVisible();
                    break;
                }
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_selectedEntry == null) return;

            var result = MessageBox.Show(
                $"スタイル「{_selectedEntry.Name}」を削除しますか？",
                "スタイル削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                _library.DeleteStyle(_selectedEntry.Name);
                _selectedEntry = null;
                RefreshList();
                UpdatePreview();
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog
            {
                Title = "スタイルライブラリのエクスポート",
                Filter = "JSONファイル (*.json)|*.json|すべてのファイル (*.*)|*.*",
                DefaultExt = "json",
                FileName = "MagosaStyleLibrary"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string json = _library.ExportToJson();
                    File.WriteAllText(dlg.FileName, json, System.Text.Encoding.UTF8);
                    ErrorHandler.ShowOperationSuccess("エクスポート",
                        $"{_library.Count}件のスタイルをエクスポートしました\n→ {dlg.FileName}");
                }
                catch (Exception ex)
                {
                    ErrorHandler.ShowOperationError("エクスポート", ex);
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "スタイルライブラリのインポート",
                Filter = "JSONファイル (*.json)|*.json|すべてのファイル (*.*)|*.*",
                DefaultExt = "json"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string json = File.ReadAllText(dlg.FileName, System.Text.Encoding.UTF8);

                    var overwrite = MessageBox.Show(
                        "同名のスタイルが存在する場合、上書きしますか？\n[はい] 上書き ／ [いいえ] スキップ",
                        "インポート", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    int count = _library.ImportFromJson(json, overwrite == DialogResult.Yes);
                    ErrorHandler.ShowOperationSuccess("インポート", $"{count}件のスタイルをインポートしました");
                    RefreshList();
                }
                catch (Exception ex)
                {
                    ErrorHandler.ShowOperationError("インポート", ex);
                }
            }
        }

        #endregion

        #region リソース管理

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _lvStyles?.Dispose();
                _btnClose?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion
    }

    /// <summary>
    /// スタイル名入力ダイアログ
    /// </summary>
    public class StyleNameInputDialog : Form
    {
        private readonly ShapeStyleLibrary _library;
        private TextBox _txtName;
        private Button _btnOk;
        private Button _btnCancel;
        private Label _lblWarning;

        public string StyleName => _txtName.Text.Trim();

        public StyleNameInputDialog(ShapeStyleLibrary library)
        {
            _library = library;
            this.Text = "スタイル名を入力";
            this.Size = new Size(380, 160);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lbl = new Label
            {
                Text = "スタイル名:",
                Location = new Point(16, 20),
                Size = new Size(80, 22),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _txtName = new TextBox
            {
                Location = new Point(100, 18),
                Size = new Size(250, 22),
                MaxLength = 50
            };
            _txtName.TextChanged += (s, e) => ValidateName();

            _lblWarning = new Label
            {
                Location = new Point(16, 46),
                Size = new Size(340, 18),
                ForeColor = Color.DarkRed,
                Text = ""
            };

            _btnOk = new Button
            {
                Text = "OK",
                Location = new Point(175, 72),
                Size = new Size(80, 26),
                DialogResult = DialogResult.OK,
                Enabled = false
            };
            _btnOk.Click += (s, e) =>
            {
                if (!ValidateName()) return;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            _btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(262, 72),
                Size = new Size(88, 26),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] { lbl, _txtName, _lblWarning, _btnOk, _btnCancel });
            this.AcceptButton = _btnOk;
            this.CancelButton = _btnCancel;
        }

        private bool ValidateName()
        {
            string name = _txtName.Text.Trim();
            if (string.IsNullOrEmpty(name))
            {
                _lblWarning.Text = "";
                _btnOk.Enabled = false;
                return false;
            }
            if (_library.ExistsName(name))
            {
                _lblWarning.Text = "⚠ 同名のスタイルが既に存在します（上書きされます）";
                _btnOk.Enabled = true;
                return true;
            }
            _lblWarning.Text = "";
            _btnOk.Enabled = true;
            return true;
        }
    }
}
