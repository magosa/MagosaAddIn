using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MagosaAddIn.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// テキスト一括編集ダイアログ
    /// 選択図形のテキストを一覧表示して一括編集・検索置換・書式統一・レイアウト調整を行う
    /// </summary>
    public class TextBulkEditDialog : BaseDialog
    {
        #region フィールド

        private readonly List<PowerPoint.Shape> _shapes;
        private readonly ShapeTextEditor _editor;
        private List<ShapeTextInfo> _textInfos;

        // タブコントロール
        private TabControl _tabControl;

        // --- テキスト編集タブ ---
        private DataGridView _dgvTexts;
        private Button _btnApplyTexts;
        private Button _btnSetUniform;
        private Button _btnDistribute;
        private Button _btnClearTexts;
        private TextBox _txtUniformInput;    // 全図形に同じテキストを設定（1行）
        private TextBox _txtDistributeInput; // テキストを配布（複数行、改行区切り）
        private Label _lblTextEditInfo;

        // --- 検索・置換タブ ---
        private TextBox _txtSearch;
        private TextBox _txtReplace;
        private CheckBox _chkCaseSensitive;
        private Button _btnSearchReplace;
        private Button _btnSearchFind;
        private Label _lblSearchResult;

        // --- フォント書式タブ ---
        private CheckBox _chkFontName;
        private ComboBox _cmbFontName;
        private CheckBox _chkFontSize;
        private NumericUpDown _numFontSize;
        private CheckBox _chkBold;
        private CheckBox _rdoBold;
        private CheckBox _chkItalic;
        private CheckBox _rdoItalic;
        private CheckBox _chkUnderline;
        private CheckBox _rdoUnderline;
        private CheckBox _chkFontColor;
        private Button _btnFontColor;
        private Panel _pnlFontColorPreview;
        private Button _btnApplyFont;
        private Color _selectedFontColor = Color.Black;

        // --- レイアウトタブ ---
        private CheckBox _chkLineSpacing;
        private NumericUpDown _numLineSpacing;
        private CheckBox _chkMarginLeft;
        private NumericUpDown _numMarginLeft;
        private CheckBox _chkMarginRight;
        private NumericUpDown _numMarginRight;
        private CheckBox _chkMarginTop;
        private NumericUpDown _numMarginTop;
        private CheckBox _chkMarginBottom;
        private NumericUpDown _numMarginBottom;
        private Button _btnApplyLayout;

        // 閉じるボタン
        private Button _btnClose;

        #endregion

        #region コンストラクタ

        public TextBulkEditDialog(List<PowerPoint.Shape> shapes)
        {
            _shapes = shapes ?? throw new ArgumentNullException(nameof(shapes));
            _editor = new ShapeTextEditor();
            _textInfos = _editor.GetTextInfos(_shapes);
            InitializeDialog();
        }

        #endregion

        #region 初期化

        private void InitializeDialog()
        {
            ConfigureForm("テキスト一括編集", 640, 560);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(640, 560);
            BuildUI();
            LoadTextData();
        }

        private void BuildUI()
        {
            _tabControl = new TabControl
            {
                Location = new Point(10, 10),
                Size = new Size(610, 460),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            var tabTextEdit = new TabPage("テキスト編集");
            var tabSearchReplace = new TabPage("検索・置換");
            var tabFont = new TabPage("フォント書式");
            var tabLayout = new TabPage("レイアウト");

            BuildTextEditTab(tabTextEdit);
            BuildSearchReplaceTab(tabSearchReplace);
            BuildFontTab(tabFont);
            BuildLayoutTab(tabLayout);

            _tabControl.TabPages.Add(tabTextEdit);
            _tabControl.TabPages.Add(tabSearchReplace);
            _tabControl.TabPages.Add(tabFont);
            _tabControl.TabPages.Add(tabLayout);

            _btnClose = new Button
            {
                Text = "閉じる",
                Size = new Size(90, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _btnClose.Location = new Point(this.ClientSize.Width - _btnClose.Width - 12,
                this.ClientSize.Height - _btnClose.Height - 12);
            _btnClose.Click += (s, e) => this.Close();

            this.Controls.Add(_tabControl);
            this.Controls.Add(_btnClose);

            this.Resize += (s, e) =>
            {
                _tabControl.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 60);
                _btnClose.Location = new Point(this.ClientSize.Width - _btnClose.Width - 12,
                    this.ClientSize.Height - _btnClose.Height - 12);
            };
        }

        #endregion

        #region テキスト編集タブ

        private void BuildTextEditTab(TabPage tab)
        {
            // 情報ラベル
            _lblTextEditInfo = new Label
            {
                Text = $"選択図形: {_shapes.Count}個（テキストフレームあり: {_textInfos.Count(t => t.HasTextFrame)}個）",
                Location = new Point(8, 8),
                Size = new Size(560, 18),
                ForeColor = Color.DarkBlue
            };

            // DataGridView（テキスト個別編集）
            _dgvTexts = new DataGridView
            {
                Location = new Point(8, 30),
                Size = new Size(580, 190),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersWidth = 30,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            _dgvTexts.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "図形名", Name = "colShapeName", ReadOnly = true, FillWeight = 30,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(245, 245, 245) }
            });
            _dgvTexts.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "テキスト内容（編集可）", Name = "colText", ReadOnly = false, FillWeight = 70
            });
            _dgvTexts.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "テキスト枠", Name = "colHasFrame", ReadOnly = true, FillWeight = 15, Width = 65
            });
            _dgvTexts.CellBeginEdit += DgvTexts_CellBeginEdit;

            // ─── グリッド操作ボタン ───
            _btnApplyTexts = new Button
            {
                Text = "グリッドの内容を適用",
                Location = new Point(8, 228),
                Size = new Size(145, 25)
            };
            new ToolTip().SetToolTip(_btnApplyTexts, "グリッドで編集したテキストを各図形に書き込みます");
            _btnApplyTexts.Click += BtnApplyTexts_Click;

            _btnClearTexts = new Button
            {
                Text = "テキストを一括削除",
                Location = new Point(160, 228),
                Size = new Size(145, 25),
                ForeColor = Color.FromArgb(180, 0, 0)
            };
            _btnClearTexts.Click += BtnClearTexts_Click;

            // ─── 区切り線 ───
            var sep1 = new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(8, 266), Size = new Size(580, 2) };

            // ─── 全図形に同じテキストを設定 ───
            var lblUniform = new Label
            {
                Text = "全図形に同じテキストを設定:",
                Location = new Point(8, 276),
                Size = new Size(185, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _txtUniformInput = new TextBox
            {
                Location = new Point(198, 274),
                Size = new Size(270, 22)
            };
            _btnSetUniform = new Button
            {
                Text = "一括設定",
                Location = new Point(476, 273),
                Size = new Size(80, 24)
            };
            _btnSetUniform.Click += BtnSetUniform_Click;

            // ─── 区切り線 ───
            var sep2 = new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(8, 308), Size = new Size(580, 2) };

            // ─── テキストを配布（専用マルチラインTextBox）───
            var lblDistribute = new Label
            {
                Text = "テキストを配布（1行 = 1図形、Enterで改行）:",
                Location = new Point(8, 317),
                Size = new Size(400, 18),
                ForeColor = Color.DarkSlateGray
            };
            _txtDistributeInput = new TextBox
            {
                Location = new Point(8, 338),
                Size = new Size(470, 75),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
                AcceptsTab = false,
                WordWrap = false
            };
            new ToolTip().SetToolTip(_txtDistributeInput,
                "例:\nりんご\nみかん\nぶどう\n→ 図形1=りんご、図形2=みかん、図形3=ぶどう");

            _btnDistribute = new Button
            {
                Text = "配布実行",
                Location = new Point(486, 338),
                Size = new Size(90, 75)
            };
            new ToolTip().SetToolTip(_btnDistribute, "入力したテキストを改行で分割し、各図形に順番に配布します");
            _btnDistribute.Click += BtnDistribute_Click;

            tab.Controls.AddRange(new Control[]
            {
                _lblTextEditInfo, _dgvTexts,
                _btnApplyTexts, _btnClearTexts,
                sep1,
                lblUniform, _txtUniformInput, _btnSetUniform,
                sep2,
                lblDistribute, _txtDistributeInput, _btnDistribute
            });
        }

        private void DgvTexts_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == _dgvTexts.Columns["colText"].Index)
            {
                bool hasFrame = (bool)(_dgvTexts.Rows[e.RowIndex].Cells["colHasFrame"].Value ?? false);
                if (!hasFrame) e.Cancel = true;
            }
        }

        private void LoadTextData()
        {
            _dgvTexts.Rows.Clear();
            foreach (var info in _textInfos)
            {
                var row = _dgvTexts.Rows.Add(info.ShapeName, info.Text, info.HasTextFrame);
                if (!info.HasTextFrame)
                {
                    _dgvTexts.Rows[row].DefaultCellStyle.ForeColor = Color.Gray;
                    _dgvTexts.Rows[row].DefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
                }
            }
        }

        /// <summary>
        /// グリッドで編集したテキストを各図形に適用
        /// </summary>
        private void BtnApplyTexts_Click(object sender, EventArgs e)
        {
            // 編集中セルを確定させる
            _dgvTexts.EndEdit();

            for (int i = 0; i < _dgvTexts.Rows.Count && i < _textInfos.Count; i++)
            {
                _textInfos[i].Text = _dgvTexts.Rows[i].Cells["colText"].Value?.ToString() ?? "";
            }

            int count = _editor.ApplyIndividualTexts(_textInfos);
            ErrorHandler.ShowOperationSuccess("テキスト適用", $"{count}個の図形にテキストを適用しました");
            _textInfos = _editor.GetTextInfos(_shapes);
            LoadTextData();
        }

        /// <summary>
        /// 全図形に同じテキストを設定
        /// </summary>
        private void BtnSetUniform_Click(object sender, EventArgs e)
        {
            string text = _txtUniformInput.Text;
            int count = _editor.SetUniformText(_shapes, text);
            ErrorHandler.ShowOperationSuccess("テキスト一括設定", $"{count}個の図形に「{text}」を設定しました");
            _textInfos = _editor.GetTextInfos(_shapes);
            LoadTextData();
        }

        /// <summary>
        /// 配布テキストボックスの内容を改行で分割して各図形に配布
        /// </summary>
        private void BtnDistribute_Click(object sender, EventArgs e)
        {
            // 専用テキストボックスから読み込む
            string sourceText = _txtDistributeInput.Text;

            if (string.IsNullOrEmpty(sourceText))
            {
                MessageBox.Show(
                    "配布するテキストを入力してください。\n\n例:\nりんご\nみかん\nぶどう",
                    "テキスト配布", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var lines = sourceText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            // 空行を除いた実際の行数を表示
            int nonEmptyLines = lines.Count(l => !string.IsNullOrEmpty(l));

            var result = MessageBox.Show(
                $"テキストを改行で分割し、{_shapes.Count}個の図形に配布します。\n\n" +
                $"入力行数: {lines.Length}行（空行除く: {nonEmptyLines}行）\n" +
                $"対象図形数: {_shapes.Count}個\n\n" +
                $"行数が図形数より少ない場合、余った図形は空になります。\n実行しますか？",
                "テキスト配布確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                int count = _editor.DistributeText(_shapes, sourceText);
                ErrorHandler.ShowOperationSuccess("テキスト配布", $"{count}個の図形にテキストを配布しました");
                _textInfos = _editor.GetTextInfos(_shapes);
                LoadTextData();
            }
        }

        /// <summary>
        /// 全図形のテキストを削除
        /// </summary>
        private void BtnClearTexts_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                $"選択した{_shapes.Count}個の図形のテキストをすべて削除します。\n実行しますか？",
                "テキスト削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                int count = _editor.ClearText(_shapes);
                ErrorHandler.ShowOperationSuccess("テキスト削除", $"{count}個の図形のテキストを削除しました");
                _textInfos = _editor.GetTextInfos(_shapes);
                LoadTextData();
            }
        }

        #endregion

        #region 検索・置換タブ

        private void BuildSearchReplaceTab(TabPage tab)
        {
            int y = 20;

            var lblSearch = new Label { Text = "検索文字列:", Location = new Point(12, y + 3), Size = new Size(90, 20) };
            _txtSearch = new TextBox { Location = new Point(108, y), Size = new Size(300, 22) };
            y += 38;

            var lblReplace = new Label { Text = "置換文字列:", Location = new Point(12, y + 3), Size = new Size(90, 20) };
            _txtReplace = new TextBox { Location = new Point(108, y), Size = new Size(300, 22) };
            y += 38;

            _chkCaseSensitive = new CheckBox
            {
                Text = "大文字/小文字を区別する",
                Location = new Point(108, y),
                Size = new Size(200, 22),
                Checked = false
            };
            y += 40;

            _btnSearchReplace = new Button
            {
                Text = "すべて置換",
                Location = new Point(108, y),
                Size = new Size(100, 28),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnSearchReplace.Click += BtnSearchReplace_Click;

            _btnSearchFind = new Button { Text = "検索のみ", Location = new Point(218, y), Size = new Size(100, 28) };
            _btnSearchFind.Click += BtnSearchFind_Click;
            y += 50;

            _lblSearchResult = new Label { Text = "", Location = new Point(108, y), Size = new Size(400, 22), ForeColor = Color.DarkBlue };
            y += 40;

            var lblHelp = new Label
            {
                Text = "※ 選択した図形内のテキストのみが対象です",
                Location = new Point(12, y),
                Size = new Size(400, 20),
                ForeColor = Color.Gray,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 8)
            };

            tab.Controls.AddRange(new Control[]
            {
                lblSearch, _txtSearch, lblReplace, _txtReplace,
                _chkCaseSensitive, _btnSearchReplace, _btnSearchFind,
                _lblSearchResult, lblHelp
            });
        }

        private void BtnSearchReplace_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtSearch.Text))
            {
                MessageBox.Show("検索文字列を入力してください。", "検索・置換", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int count = _editor.SearchAndReplace(_shapes, _txtSearch.Text, _txtReplace.Text, _chkCaseSensitive.Checked);
            _lblSearchResult.Text = count > 0 ? $"✓ {count}個の図形で置換しました" : "一致するテキストが見つかりませんでした";
            _lblSearchResult.ForeColor = count > 0 ? Color.DarkGreen : Color.DarkRed;
            if (count > 0) { _textInfos = _editor.GetTextInfos(_shapes); LoadTextData(); }
        }

        private void BtnSearchFind_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtSearch.Text))
            {
                MessageBox.Show("検索文字列を入力してください。", "テキスト検索", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var found = _editor.FindShapesByText(_shapes, _txtSearch.Text, _chkCaseSensitive.Checked);
            _lblSearchResult.Text = found.Count > 0
                ? $"✓ {found.Count}個の図形で「{_txtSearch.Text}」が見つかりました"
                : $"「{_txtSearch.Text}」は見つかりませんでした";
            _lblSearchResult.ForeColor = found.Count > 0 ? Color.DarkGreen : Color.DarkRed;
        }

        #endregion

        #region フォント書式タブ

        private void BuildFontTab(TabPage tab)
        {
            int y = 14;
            int controlX = 140;

            _chkFontName = new CheckBox { Text = "フォント名:", Location = new Point(10, y + 2), Size = new Size(125, 20) };
            _cmbFontName = new ComboBox
            {
                Location = new Point(controlX, y),
                Size = new Size(220, 22),
                DropDownStyle = ComboBoxStyle.DropDown,
                Enabled = false,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            // PowerPoint と同じフォント一覧: システムにインストール済みの全フォントをアルファベット順で追加
            var allFontNames = System.Drawing.FontFamily.Families
                .Select(f => f.Name)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            _cmbFontName.Items.AddRange(allFontNames);
            if (_cmbFontName.Items.Count > 0) _cmbFontName.SelectedIndex = 0;
            _chkFontName.CheckedChanged += (s, e) => _cmbFontName.Enabled = _chkFontName.Checked;
            y += 34;

            _chkFontSize = new CheckBox { Text = "フォントサイズ:", Location = new Point(10, y + 2), Size = new Size(125, 20) };
            _numFontSize = new NumericUpDown { Location = new Point(controlX, y), Size = new Size(80, 22), Minimum = 1, Maximum = 200, Value = 18, Enabled = false };
            var lblPt = new Label { Text = "pt", Location = new Point(controlX + 85, y + 3), Size = new Size(25, 20) };
            _chkFontSize.CheckedChanged += (s, e) => _numFontSize.Enabled = _chkFontSize.Checked;
            y += 34;

            _chkBold = new CheckBox { Text = "太字:", Location = new Point(10, y + 2), Size = new Size(65, 20) };
            _rdoBold = new CheckBox { Text = "適用する", Location = new Point(controlX, y + 2), Size = new Size(90, 20), Checked = true, Enabled = false };
            _chkBold.CheckedChanged += (s, e) => _rdoBold.Enabled = _chkBold.Checked;
            y += 30;

            _chkItalic = new CheckBox { Text = "斜体:", Location = new Point(10, y + 2), Size = new Size(65, 20) };
            _rdoItalic = new CheckBox { Text = "適用する", Location = new Point(controlX, y + 2), Size = new Size(90, 20), Checked = true, Enabled = false };
            _chkItalic.CheckedChanged += (s, e) => _rdoItalic.Enabled = _chkItalic.Checked;
            y += 30;

            _chkUnderline = new CheckBox { Text = "下線:", Location = new Point(10, y + 2), Size = new Size(65, 20) };
            _rdoUnderline = new CheckBox { Text = "適用する", Location = new Point(controlX, y + 2), Size = new Size(90, 20), Checked = true, Enabled = false };
            _chkUnderline.CheckedChanged += (s, e) => _rdoUnderline.Enabled = _chkUnderline.Checked;
            y += 34;

            _chkFontColor = new CheckBox { Text = "フォント色:", Location = new Point(10, y + 2), Size = new Size(95, 20) };
            _pnlFontColorPreview = new Panel { Location = new Point(controlX, y + 1), Size = new Size(22, 22), BackColor = Color.Black, BorderStyle = BorderStyle.FixedSingle, Enabled = false };
            _btnFontColor = new Button { Text = "色を選択...", Location = new Point(controlX + 28, y), Size = new Size(90, 24), Enabled = false };
            _btnFontColor.Click += BtnFontColor_Click;
            _chkFontColor.CheckedChanged += (s, e) => { _pnlFontColorPreview.Enabled = _chkFontColor.Checked; _btnFontColor.Enabled = _chkFontColor.Checked; };
            y += 44;

            var sep = new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(10, y), Size = new Size(560, 2) };
            y += 10;

            _btnApplyFont = new Button
            {
                Text = "フォント書式を適用",
                Location = new Point(10, y),
                Size = new Size(145, 28),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnApplyFont.Click += BtnApplyFont_Click;

            tab.Controls.AddRange(new Control[]
            {
                _chkFontName, _cmbFontName,
                _chkFontSize, _numFontSize, lblPt,
                _chkBold, _rdoBold,
                _chkItalic, _rdoItalic,
                _chkUnderline, _rdoUnderline,
                _chkFontColor, _pnlFontColorPreview, _btnFontColor,
                sep, _btnApplyFont
            });
        }

        private void BtnFontColor_Click(object sender, EventArgs e)
        {
            using (var dlg = new ColorDialog { Color = _selectedFontColor, FullOpen = true })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _selectedFontColor = dlg.Color;
                    _pnlFontColorPreview.BackColor = _selectedFontColor;
                }
            }
        }

        private void BtnApplyFont_Click(object sender, EventArgs e)
        {
            var settings = new FontSettings();
            bool anyChecked = false;

            if (_chkFontName.Checked && _cmbFontName.SelectedItem != null) { settings.FontName = _cmbFontName.SelectedItem.ToString(); anyChecked = true; }
            if (_chkFontSize.Checked) { settings.FontSize = (float)_numFontSize.Value; anyChecked = true; }
            if (_chkBold.Checked) { settings.IsBold = _rdoBold.Checked; anyChecked = true; }
            if (_chkItalic.Checked) { settings.IsItalic = _rdoItalic.Checked; anyChecked = true; }
            if (_chkUnderline.Checked) { settings.IsUnderline = _rdoUnderline.Checked; anyChecked = true; }
            if (_chkFontColor.Checked)
            {
                settings.FontColor = _selectedFontColor.R | (_selectedFontColor.G << 8) | (_selectedFontColor.B << 16);
                anyChecked = true;
            }

            if (!anyChecked) { MessageBox.Show("変更する項目をチェックしてください。", "フォント書式", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            int count = _editor.ApplyFontSettings(_shapes, settings);
            ErrorHandler.ShowOperationSuccess("フォント書式適用", $"{count}個の図形にフォント書式を適用しました");
        }

        #endregion

        #region レイアウトタブ

        private void BuildLayoutTab(TabPage tab)
        {
            int y = 14;
            int controlX = 160;

            var lblHelp = new Label
            {
                Text = "チェックした項目のみ変更されます（pt単位）",
                Location = new Point(10, y), Size = new Size(400, 18),
                ForeColor = Color.Gray, Font = new Font(SystemFonts.DefaultFont.FontFamily, 8)
            };
            y += 30;

            _chkLineSpacing = new CheckBox { Text = "行間:", Location = new Point(10, y + 2), Size = new Size(65, 20) };
            _numLineSpacing = new NumericUpDown { Location = new Point(controlX, y), Size = new Size(80, 22), Minimum = (decimal)0.5, Maximum = 10, Value = 1, DecimalPlaces = 1, Increment = (decimal)0.1, Enabled = false };
            var lblLS = new Label { Text = "倍", Location = new Point(controlX + 85, y + 3), Size = new Size(30, 20) };
            _chkLineSpacing.CheckedChanged += (s, e) => _numLineSpacing.Enabled = _chkLineSpacing.Checked;
            y += 34;

            _chkMarginLeft = new CheckBox { Text = "左余白:", Location = new Point(10, y + 2), Size = new Size(80, 20) };
            _numMarginLeft = new NumericUpDown { Location = new Point(controlX, y), Size = new Size(80, 22), Minimum = 0, Maximum = 100, Value = (decimal)7.2, DecimalPlaces = 1, Increment = (decimal)0.5, Enabled = false };
            var lblML = new Label { Text = "pt", Location = new Point(controlX + 85, y + 3), Size = new Size(25, 20) };
            _chkMarginLeft.CheckedChanged += (s, e) => _numMarginLeft.Enabled = _chkMarginLeft.Checked;
            y += 34;

            _chkMarginRight = new CheckBox { Text = "右余白:", Location = new Point(10, y + 2), Size = new Size(80, 20) };
            _numMarginRight = new NumericUpDown { Location = new Point(controlX, y), Size = new Size(80, 22), Minimum = 0, Maximum = 100, Value = (decimal)7.2, DecimalPlaces = 1, Increment = (decimal)0.5, Enabled = false };
            var lblMR = new Label { Text = "pt", Location = new Point(controlX + 85, y + 3), Size = new Size(25, 20) };
            _chkMarginRight.CheckedChanged += (s, e) => _numMarginRight.Enabled = _chkMarginRight.Checked;
            y += 34;

            _chkMarginTop = new CheckBox { Text = "上余白:", Location = new Point(10, y + 2), Size = new Size(80, 20) };
            _numMarginTop = new NumericUpDown { Location = new Point(controlX, y), Size = new Size(80, 22), Minimum = 0, Maximum = 100, Value = (decimal)3.6, DecimalPlaces = 1, Increment = (decimal)0.5, Enabled = false };
            var lblMT = new Label { Text = "pt", Location = new Point(controlX + 85, y + 3), Size = new Size(25, 20) };
            _chkMarginTop.CheckedChanged += (s, e) => _numMarginTop.Enabled = _chkMarginTop.Checked;
            y += 34;

            _chkMarginBottom = new CheckBox { Text = "下余白:", Location = new Point(10, y + 2), Size = new Size(80, 20) };
            _numMarginBottom = new NumericUpDown { Location = new Point(controlX, y), Size = new Size(80, 22), Minimum = 0, Maximum = 100, Value = (decimal)3.6, DecimalPlaces = 1, Increment = (decimal)0.5, Enabled = false };
            var lblMB = new Label { Text = "pt", Location = new Point(controlX + 85, y + 3), Size = new Size(25, 20) };
            _chkMarginBottom.CheckedChanged += (s, e) => _numMarginBottom.Enabled = _chkMarginBottom.Checked;
            y += 44;

            var sep = new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(10, y), Size = new Size(540, 2) };
            y += 10;

            _btnApplyLayout = new Button
            {
                Text = "レイアウトを適用",
                Location = new Point(10, y),
                Size = new Size(130, 28),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnApplyLayout.Click += BtnApplyLayout_Click;

            tab.Controls.AddRange(new Control[]
            {
                lblHelp,
                _chkLineSpacing, _numLineSpacing, lblLS,
                _chkMarginLeft, _numMarginLeft, lblML,
                _chkMarginRight, _numMarginRight, lblMR,
                _chkMarginTop, _numMarginTop, lblMT,
                _chkMarginBottom, _numMarginBottom, lblMB,
                sep, _btnApplyLayout
            });
        }

        private void BtnApplyLayout_Click(object sender, EventArgs e)
        {
            var settings = new TextLayoutSettings();
            bool anyChecked = false;

            if (_chkLineSpacing.Checked) { settings.LineSpacingPt = (float)_numLineSpacing.Value; anyChecked = true; }
            if (_chkMarginLeft.Checked) { settings.MarginLeft = (float)_numMarginLeft.Value; anyChecked = true; }
            if (_chkMarginRight.Checked) { settings.MarginRight = (float)_numMarginRight.Value; anyChecked = true; }
            if (_chkMarginTop.Checked) { settings.MarginTop = (float)_numMarginTop.Value; anyChecked = true; }
            if (_chkMarginBottom.Checked) { settings.MarginBottom = (float)_numMarginBottom.Value; anyChecked = true; }

            if (!anyChecked) { MessageBox.Show("変更する項目をチェックしてください。", "レイアウト設定", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            int count = _editor.ApplyTextLayoutSettings(_shapes, settings);
            ErrorHandler.ShowOperationSuccess("レイアウト適用", $"{count}個の図形にレイアウト設定を適用しました");
        }

        #endregion

        #region リソース管理

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _tabControl?.Dispose();
                _btnClose?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
