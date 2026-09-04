using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MagosaAddIn.Core;
using ColorConv = MagosaAddIn.Core.ColorConverter;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// スライド内シェイプの色（塗り・線・フォント）を収集し、一括置換するダイアログ
    /// </summary>
    public partial class ColorReplaceDialog : BaseDialog
    {
        #region フィールド

        private readonly ColorReplacer _replacer = new ColorReplacer();
        private readonly ColorReplaceLibrary _library = new ColorReplaceLibrary();

        private RadioButton _rbCurrentSlide;
        private RadioButton _rbAllSlides;
        private Button _btnCollect;
        private ListView _lvColors;
        private Button _btnSetReplacement;
        private Button _btnClearReplacement;
        private TextBox _txtHexReplacement;
        private ComboBox _cmbSavedLists;
        private Button _btnLoadList;
        private Button _btnDeleteList;
        private Button _btnSaveList;
        private Label _lblStatus;
        private Button _btnApply;
        private Button _btnClose;

        private bool _hasCollected;

        #endregion

        #region 内部クラス

        /// <summary>1色分の収集結果と設定中の置換色を保持する行データ</summary>
        private class ColorReplaceRow
        {
            public int OriginalColor;
            public int UsageCount;
            public int? ReplacementColor;
        }

        #endregion

        #region コンストラクタ

        public ColorReplaceDialog()
        {
            InitializeDialog();
            RefreshSavedListCombo();
        }

        #endregion

        private ColorReplaceScope CurrentScope =>
            _rbAllSlides.Checked ? ColorReplaceScope.AllSlides : ColorReplaceScope.CurrentSlide;

        #region UI初期化

        private void InitializeDialog()
        {
            ConfigureForm("色置換", 680, 580);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(680, 480);
            BuildUI();
        }

        private void BuildUI()
        {
            // ===== 対象範囲 =====
            var grpScope = CreateGroupBox("対象範囲", new Point(20, 20), new Size(640, 55));
            grpScope.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            _rbCurrentSlide = CreateRadioButton("現在のスライドのみ", new Point(15, 18), new Size(160, 20), isChecked: true);
            _rbAllSlides = CreateRadioButton("プレゼンテーション全体", new Point(185, 18), new Size(190, 20));
            _btnCollect = new Button { Text = "色を取得", Location = new Point(500, 14), Size = new Size(120, 28) };
            _btnCollect.Click += BtnCollect_Click;

            grpScope.Controls.AddRange(new Control[] { _rbCurrentSlide, _rbAllSlides, _btnCollect });

            // ===== 色一覧 =====
            _lvColors = new ListView
            {
                Location = new Point(20, 87),
                Size = new Size(640, 260),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = true,
                HideSelection = false,
                OwnerDraw = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _lvColors.SmallImageList = new ImageList { ImageSize = new Size(1, 20) }; // 行の高さをスウォッチが見える大きさに広げるため
            _lvColors.Columns.Add("色", 50, HorizontalAlignment.Center);
            _lvColors.Columns.Add("カラーコード", 100, HorizontalAlignment.Left);
            _lvColors.Columns.Add("使用数", 60, HorizontalAlignment.Center);
            _lvColors.Columns.Add("置換後", 50, HorizontalAlignment.Center);
            _lvColors.Columns.Add("置換後カラーコード", 130, HorizontalAlignment.Left);
            _lvColors.DrawColumnHeader += (s, e) => e.DrawDefault = true;
            _lvColors.DrawItem += (s, e) => { }; // true にすると Details 表示で DrawSubItem が呼ばれなくなるため何もしない
            _lvColors.DrawSubItem += LvColors_DrawSubItem;
            _lvColors.SelectedIndexChanged += (s, e) => UpdateButtonStates();

            // ===== 置換操作ボタン =====
            _btnSetReplacement = new Button
            {
                Text = "置換色を設定...",
                Location = new Point(20, 357),
                Size = new Size(180, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _btnSetReplacement.Click += BtnSetReplacement_Click;

            _btnClearReplacement = new Button
            {
                Text = "変更なしに戻す",
                Location = new Point(210, 357),
                Size = new Size(150, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _btnClearReplacement.Click += BtnClearReplacement_Click;

            var lblHex = CreateLabel("HEX指定:", new Point(375, 361), 60);
            lblHex.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _txtHexReplacement = new TextBox
            {
                Location = new Point(440, 360),
                Size = new Size(90, 22),
                MaxLength = 7,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _txtHexReplacement.KeyDown += TxtHexReplacement_KeyDown;
            _txtHexReplacement.Leave += (s, e) => ApplyHexReplacement();

            var separator = new Label
            {
                BorderStyle = BorderStyle.Fixed3D,
                Location = new Point(20, 397),
                Size = new Size(640, 2),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            // ===== 保存済みリスト =====
            var grpSaved = CreateGroupBox("保存済みリスト", new Point(20, 407), new Size(640, 70));
            grpSaved.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            var lblSaved = CreateLabel("リスト名:", new Point(15, 27), 70);
            _cmbSavedLists = CreateComboBox(new Point(90, 26), new Size(240, 22));
            _btnLoadList = new Button { Text = "読込", Location = new Point(340, 23), Size = new Size(70, 28) };
            _btnLoadList.Click += BtnLoadList_Click;
            _btnDeleteList = new Button { Text = "削除", Location = new Point(415, 23), Size = new Size(70, 28) };
            _btnDeleteList.Click += BtnDeleteList_Click;
            _btnSaveList = new Button { Text = "名前を付けて保存...", Location = new Point(495, 23), Size = new Size(130, 28) };
            _btnSaveList.Click += BtnSaveList_Click;

            grpSaved.Controls.AddRange(new Control[] { lblSaved, _cmbSavedLists, _btnLoadList, _btnDeleteList, _btnSaveList });

            // ===== ステータス・適用・閉じる =====
            _lblStatus = new Label
            {
                Location = new Point(20, 487),
                Size = new Size(440, 28),
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _btnApply = new Button
            {
                Text = "適用",
                Location = new Point(470, 487),
                Size = new Size(90, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _btnApply.Click += BtnApply_Click;
            _btnClose = new Button
            {
                Text = "閉じる",
                Location = new Point(570, 487),
                Size = new Size(90, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[]
            {
                grpScope, _lvColors, _btnSetReplacement, _btnClearReplacement, lblHex, _txtHexReplacement, separator,
                grpSaved, _lblStatus, _btnApply, _btnClose
            });

            UpdateButtonStates();
        }

        #endregion

        #region 色一覧の描画

        private void LvColors_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            if (e.ColumnIndex != 0 && e.ColumnIndex != 3)
            {
                e.DrawDefault = true;
                return;
            }

            var row = e.Item.Tag as ColorReplaceRow;
            if (row == null) { e.DrawDefault = true; return; }

            e.Graphics.FillRectangle(e.Item.Selected ? SystemBrushes.Highlight : SystemBrushes.Window, e.Bounds);

            int? rgb = e.ColumnIndex == 0 ? row.OriginalColor : row.ReplacementColor;
            if (!rgb.HasValue)
            {
                return; // 置換後が未設定の場合はスウォッチなし（カラーコード側の列に「変更なし」を表示）
            }

            var rect = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 4, e.Bounds.Width - 12, e.Bounds.Height - 8);
            using (var brush = new SolidBrush(ColorConv.RgbToColor(rgb.Value)))
                e.Graphics.FillRectangle(brush, rect);
            e.Graphics.DrawRectangle(Pens.Gray, rect);
        }

        #endregion

        #region ボタンハンドラ

        private void BtnCollect_Click(object sender, EventArgs e)
        {
            if (_hasCollected && HasAnyReplacementSet())
            {
                var confirm = MessageBox.Show(
                    "設定済みの置換色がリセットされます。よろしいですか？",
                    "色の再取得", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes) return;
            }

            try
            {
                var colors = _replacer.CollectUsedColors(CurrentScope);
                PopulateListView(colors);
                _hasCollected = true;
                _lblStatus.Text = $"{colors.Count}色を検出しました。";
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("色の取得", ex);
            }
        }

        private void BtnSetReplacement_Click(object sender, EventArgs e)
        {
            if (_lvColors.SelectedItems.Count == 0) return;

            var firstRow = (ColorReplaceRow)_lvColors.SelectedItems[0].Tag;
            int initialColor = firstRow.ReplacementColor ?? firstRow.OriginalColor;

            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = ColorConv.RgbToColor(initialColor);
                colorDialog.FullOpen = true;

                if (colorDialog.ShowDialog() != DialogResult.OK) return;

                ApplyReplacementColor(ColorConv.ColorToRgb(colorDialog.Color));
            }
        }

        private void TxtHexReplacement_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                ApplyHexReplacement();
            }
        }

        /// <summary>HEXテキストボックスの入力を選択中の行の置換色に反映する</summary>
        private void ApplyHexReplacement()
        {
            string hex = _txtHexReplacement.Text.Trim();
            if (hex.Length == 0) return;
            if (!hex.StartsWith("#")) hex = "#" + hex;

            if (_lvColors.SelectedItems.Count == 0 || hex.Length != 7) return;

            try
            {
                int newColor = ColorConv.HexToRgb(hex);
                ApplyReplacementColor(newColor);
            }
            catch
            {
                // 無効なHEX入力は無視
            }
        }

        /// <summary>選択中の行すべての置換色を指定色に設定する</summary>
        private void ApplyReplacementColor(int newColor)
        {
            foreach (ListViewItem item in _lvColors.SelectedItems)
            {
                ((ColorReplaceRow)item.Tag).ReplacementColor = newColor;
                item.SubItems[4].Text = ColorConv.RgbToHex(newColor);
            }

            _txtHexReplacement.Text = ColorConv.RgbToHex(newColor);
            _lvColors.Invalidate();
            _lvColors.Update();
        }

        private void BtnClearReplacement_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in _lvColors.SelectedItems)
            {
                ((ColorReplaceRow)item.Tag).ReplacementColor = null;
                item.SubItems[4].Text = "(変更なし)";
            }
            _lvColors.Invalidate();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            var map = BuildReplacementMap();
            if (map.Count == 0)
            {
                MessageBox.Show("置換する色が設定されていません。", "色置換", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var (shapeCount, propertyCount) = _replacer.ApplyReplacements(CurrentScope, map);
                string message = $"{shapeCount}個のシェイプ、{propertyCount}箇所の色を置換しました。";
                _lblStatus.Text = message;
                ErrorHandler.ShowOperationSuccess("色置換", message);
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowOperationError("色置換", ex);
            }
        }

        private void BtnSaveList_Click(object sender, EventArgs e)
        {
            if (_lvColors.Items.Count == 0) return;

            var entries = _lvColors.Items.Cast<ListViewItem>()
                .Select(i => (ColorReplaceRow)i.Tag)
                .Select(r => new ColorReplacementEntry { OriginalColor = r.OriginalColor, ReplacementColor = r.ReplacementColor })
                .ToList();

            using (var nameDialog = new ColorListNameInputDialog(_library))
            {
                if (nameDialog.ShowDialog() != DialogResult.OK) return;

                try
                {
                    _library.SaveList(nameDialog.ListName, entries);
                    RefreshSavedListCombo();
                    _lblStatus.Text = $"リスト「{nameDialog.ListName}」を保存しました。";
                }
                catch (Exception ex)
                {
                    ErrorHandler.ShowOperationError("色置換リスト保存", ex);
                }
            }
        }

        private void BtnLoadList_Click(object sender, EventArgs e)
        {
            string name = _cmbSavedLists.SelectedItem as string;
            if (string.IsNullOrEmpty(name)) return;

            var entries = _library.LoadList(name);
            if (entries == null) return;

            var map = entries
                .Where(entry => entry.ReplacementColor.HasValue)
                .ToDictionary(entry => entry.OriginalColor, entry => entry.ReplacementColor.Value);

            int applied = 0;
            foreach (ListViewItem item in _lvColors.Items)
            {
                var row = (ColorReplaceRow)item.Tag;
                if (map.TryGetValue(row.OriginalColor, out int replacement))
                {
                    row.ReplacementColor = replacement;
                    item.SubItems[4].Text = ColorConv.RgbToHex(replacement);
                    applied++;
                }
            }
            _lvColors.Invalidate();
            _lblStatus.Text = $"リスト「{name}」から{applied}件の置換設定を反映しました。";
        }

        private void BtnDeleteList_Click(object sender, EventArgs e)
        {
            string name = _cmbSavedLists.SelectedItem as string;
            if (string.IsNullOrEmpty(name)) return;

            var confirm = MessageBox.Show($"リスト「{name}」を削除しますか？",
                "色置換リスト削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            _library.DeleteList(name);
            RefreshSavedListCombo();
            _lblStatus.Text = $"リスト「{name}」を削除しました。";
        }

        #endregion

        #region 補助メソッド

        private void PopulateListView(List<ColorUsageInfo> colors)
        {
            _lvColors.BeginUpdate();
            _lvColors.Items.Clear();
            foreach (var info in colors)
            {
                var row = new ColorReplaceRow { OriginalColor = info.RgbColor, UsageCount = info.UsageCount, ReplacementColor = null };
                var item = new ListViewItem(""); // 色スウォッチは独自描画
                item.SubItems.Add(ColorConv.RgbToHex(info.RgbColor));
                item.SubItems.Add(info.UsageCount.ToString());
                item.SubItems.Add(""); // 置換後スウォッチも独自描画
                item.SubItems.Add("(変更なし)");
                item.Tag = row;
                _lvColors.Items.Add(item);
            }
            _lvColors.EndUpdate();
            UpdateButtonStates();
        }

        private bool HasAnyReplacementSet()
        {
            return _lvColors.Items.Cast<ListViewItem>()
                .Any(i => ((ColorReplaceRow)i.Tag).ReplacementColor.HasValue);
        }

        private Dictionary<int, int> BuildReplacementMap()
        {
            var map = new Dictionary<int, int>();
            foreach (ListViewItem item in _lvColors.Items)
            {
                var row = (ColorReplaceRow)item.Tag;
                if (row.ReplacementColor.HasValue && row.ReplacementColor.Value != row.OriginalColor)
                {
                    map[row.OriginalColor] = row.ReplacementColor.Value;
                }
            }
            return map;
        }

        private void RefreshSavedListCombo()
        {
            _cmbSavedLists.Items.Clear();
            foreach (var list in _library.GetAllLists().OrderBy(l => l.Name))
            {
                _cmbSavedLists.Items.Add(list.Name);
            }
            bool hasLists = _cmbSavedLists.Items.Count > 0;
            _btnLoadList.Enabled = hasLists;
            _btnDeleteList.Enabled = hasLists;
            if (hasLists) _cmbSavedLists.SelectedIndex = 0;
        }

        private void UpdateButtonStates()
        {
            bool hasSelection = _lvColors.SelectedItems.Count > 0;
            _btnSetReplacement.Enabled = hasSelection;
            _btnClearReplacement.Enabled = hasSelection;
            _txtHexReplacement.Enabled = hasSelection;

            _txtHexReplacement.Text = hasSelection
                ? ColorConv.RgbToHex(((ColorReplaceRow)_lvColors.SelectedItems[0].Tag).ReplacementColor
                    ?? ((ColorReplaceRow)_lvColors.SelectedItems[0].Tag).OriginalColor)
                : string.Empty;

            bool hasRows = _lvColors.Items.Count > 0;
            _btnApply.Enabled = hasRows;
            _btnSaveList.Enabled = hasRows;
        }

        #endregion
    }

    /// <summary>
    /// 色置換リスト名入力ダイアログ
    /// </summary>
    public class ColorListNameInputDialog : Form
    {
        private readonly ColorReplaceLibrary _library;
        private TextBox _txtName;
        private Button _btnOk;
        private Button _btnCancel;
        private Label _lblWarning;

        public string ListName => _txtName.Text.Trim();

        public ColorListNameInputDialog(ColorReplaceLibrary library)
        {
            _library = library;
            this.Text = "リスト名を入力";
            this.Size = new Size(380, 160);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lbl = new Label
            {
                Text = "リスト名:",
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
                _lblWarning.Text = "⚠ 同名のリストが既に存在します（上書きされます）";
                _btnOk.Enabled = true;
                return true;
            }
            _lblWarning.Text = "";
            _btnOk.Enabled = true;
            return true;
        }
    }
}
