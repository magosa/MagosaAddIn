using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 選択図形の処理順序を変更するダイアログ
    /// </summary>
    public partial class SelectionOrderDialog : BaseDialog
    {
        #region 公開プロパティ

        /// <summary>後続アクションの種類</summary>
        public enum PostAction
        {
            /// <summary>指定した順序で選び直すのみ（Z順は変更しない）</summary>
            ReSelectOnly,
            /// <summary>指定した順序で選び直し、さらにスタックに保存する</summary>
            ReSelectAndSaveToStack
        }

        /// <summary>OK押下後の並び替え済み図形リスト</summary>
        public List<PowerPoint.Shape> OrderedShapes { get; private set; }

        /// <summary>選択された後続アクション</summary>
        public PostAction SelectedAction { get; private set; }

        #endregion

        #region プライベートフィールド

        private readonly List<PowerPoint.Shape> _initialShapes;

        private ListView listViewShapes;
        private NoFocusButton btnMoveUp;
        private NoFocusButton btnMoveDown;
        private RadioButton rbReSelectOnly;
        private RadioButton rbReSelectAndSave;
        private PictureBox pboxThumb;
        private Label lblThumbName;

        /// <summary>図形名をキーにしたサムネイルキャッシュ（同一図形は1回だけExportする）</summary>
        private readonly Dictionary<string, Image> _thumbnailCache = new Dictionary<string, Image>();

        /// <summary>デバウンス用タイマー（連打中はExportを遅延させる）</summary>
        private System.Windows.Forms.Timer _thumbDebounceTimer;

        /// <summary>デバウンス待ち中の選択インデックス</summary>
        private int _pendingThumbIndex = -1;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// 選択順序変更ダイアログを初期化する
        /// </summary>
        /// <param name="shapes">現在選択中の図形リスト（Z順）</param>
        public SelectionOrderDialog(List<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count < 2)
                throw new ArgumentException("2個以上の図形が必要です。");

            _initialShapes = shapes;
            SelectedAction = PostAction.ReSelectOnly;
            OrderedShapes = new List<PowerPoint.Shape>(shapes);

            _thumbDebounceTimer = new System.Windows.Forms.Timer { Interval = 150 };
            _thumbDebounceTimer.Tick += ThumbDebounceTimer_Tick;

            InitializeComponent();
            PopulateList();
        }

        #endregion

        #region UI初期化

        private void InitializeComponent()
        {
            this.SuspendLayout();

            const int formWidth = 630;
            const int margin = DefaultMargin;
            const int sectionSpacing = 12;
            const int leftGroupWidth = 375;
            const int rightGroupWidth = 205;
            const int topGroupHeight = 230;
            int currentY = margin;

            ConfigureForm("選択図形の順序変更", formWidth, 460);

            // ===== 図形一覧グループ（左側）=====
            var grpList = CreateGroupBox(
                $"図形一覧（{_initialShapes.Count}個）",
                new Point(margin, currentY),
                new Size(leftGroupWidth, topGroupHeight));

            listViewShapes = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false,
                HideSelection = false,   // フォーカスが外れてもハイライトを維持
                Location = new Point(8, 20),
                Size = new Size(leftGroupWidth - 16, 160)
            };
            listViewShapes.Columns.Add("#", 36, HorizontalAlignment.Center);
            listViewShapes.Columns.Add("図形名", 170, HorizontalAlignment.Left);
            listViewShapes.Columns.Add("種類", 90, HorizontalAlignment.Left);
            listViewShapes.Columns.Add("Z順", 60, HorizontalAlignment.Center);
            listViewShapes.SelectedIndexChanged += ListViewShapes_SelectedIndexChanged;

            btnMoveUp = new NoFocusButton
            {
                Text = "↑ 上へ",
                Location = new Point(8, 188),
                Size = new Size(90, 25),
                Enabled = false
            };
            btnMoveUp.Click += BtnMoveUp_Click;

            btnMoveDown = new NoFocusButton
            {
                Text = "↓ 下へ",
                Location = new Point(104, 188),
                Size = new Size(90, 25),
                Enabled = false
            };
            btnMoveDown.Click += BtnMoveDown_Click;

            grpList.Controls.Add(listViewShapes);
            grpList.Controls.Add(btnMoveUp);
            grpList.Controls.Add(btnMoveDown);
            this.Controls.Add(grpList);

            // ===== サムネイルグループ（右側）=====
            var grpThumb = CreateGroupBox(
                "選択中の図形",
                new Point(margin + leftGroupWidth + 10, currentY),
                new Size(rightGroupWidth, topGroupHeight));

            pboxThumb = new PictureBox
            {
                Location = new Point(8, 20),
                Size = new Size(rightGroupWidth - 16, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke
            };

            lblThumbName = new Label
            {
                Location = new Point(8, 195),
                Size = new Size(rightGroupWidth - 16, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Text = "（未選択）",
                Font = SmallFont
            };

            grpThumb.Controls.Add(pboxThumb);
            grpThumb.Controls.Add(lblThumbName);
            this.Controls.Add(grpThumb);

            currentY += topGroupHeight + sectionSpacing;

            // ===== 後続アクション グループ =====
            var grpAction = CreateGroupBox(
                "後続アクション（OKを押すと必ず指定順序で選び直します）",
                new Point(margin, currentY),
                new Size(formWidth - margin * 2, 90));

            rbReSelectOnly = CreateRadioButton(
                "この順序で選び直す（Z順・体裁は変更しない）",
                new Point(12, 22), new Size(560, 20), true);

            rbReSelectAndSave = CreateRadioButton(
                "この順序で選び直し、スタックにも保存する",
                new Point(12, 50), new Size(560, 20), false);

            grpAction.Controls.AddRange(new Control[] { rbReSelectOnly, rbReSelectAndSave });
            this.Controls.Add(grpAction);

            currentY += grpAction.Height + sectionSpacing;

            // ===== OK / キャンセルボタン =====
            AddStandardButtons(currentY, BtnOK_Click);

            // フォーム高さを確定
            this.ClientSize = new Size(formWidth, CalculateFormHeight(currentY));

            this.ResumeLayout(false);
        }

        #endregion

        #region リスト操作

        /// <summary>図形リストをListViewに反映する</summary>
        private void PopulateList()
        {
            listViewShapes.Items.Clear();

            for (int i = 0; i < OrderedShapes.Count; i++)
            {
                var shape = OrderedShapes[i];
                string typeName = ShapeOrderManager.GetShapeTypeName(shape);
                int zOrder = 0;
                try { zOrder = shape.ZOrderPosition; } catch { }

                var item = new ListViewItem((i + 1).ToString());
                item.SubItems.Add(shape.Name);
                item.SubItems.Add(typeName);
                item.SubItems.Add(zOrder.ToString());
                item.Tag = shape;
                listViewShapes.Items.Add(item);
            }
        }

        /// <summary>番号列（#）のみ再採番する（リスト内容は再構築しない）</summary>
        private void RenumberItems()
        {
            for (int i = 0; i < listViewShapes.Items.Count; i++)
            {
                listViewShapes.Items[i].Text = (i + 1).ToString();
            }
        }

        /// <summary>上移動ボタン</summary>
        private void BtnMoveUp_Click(object sender, EventArgs e)
        {
            int idx = listViewShapes.SelectedIndices.Count > 0
                ? listViewShapes.SelectedIndices[0] : -1;
            if (idx <= 0) return;

            // 先にフォーカスを戻してからアイテム操作することで色抜けを防ぐ
            listViewShapes.Focus();
            listViewShapes.BeginUpdate();
            try
            {
                SwapItems(idx, idx - 1);
                listViewShapes.Items[idx - 1].Selected = true;
                listViewShapes.Items[idx - 1].EnsureVisible();
                UpdateMoveButtons(idx - 1);
            }
            finally
            {
                listViewShapes.EndUpdate();
            }
        }

        /// <summary>下移動ボタン</summary>
        private void BtnMoveDown_Click(object sender, EventArgs e)
        {
            int idx = listViewShapes.SelectedIndices.Count > 0
                ? listViewShapes.SelectedIndices[0] : -1;
            if (idx < 0 || idx >= listViewShapes.Items.Count - 1) return;

            // 先にフォーカスを戻してからアイテム操作することで色抜けを防ぐ
            listViewShapes.Focus();
            listViewShapes.BeginUpdate();
            try
            {
                SwapItems(idx, idx + 1);
                listViewShapes.Items[idx + 1].Selected = true;
                listViewShapes.Items[idx + 1].EnsureVisible();
                UpdateMoveButtons(idx + 1);
            }
            finally
            {
                listViewShapes.EndUpdate();
            }
        }

        /// <summary>指定インデックスの2行を入れ替える</summary>
        private void SwapItems(int i, int j)
        {
            var items = listViewShapes.Items;

            // サブアイテムテキストを退避
            string nameI = items[i].SubItems[1].Text;
            string typeI = items[i].SubItems[2].Text;
            string zI    = items[i].SubItems[3].Text;
            object tagI  = items[i].Tag;

            string nameJ = items[j].SubItems[1].Text;
            string typeJ = items[j].SubItems[2].Text;
            string zJ    = items[j].SubItems[3].Text;
            object tagJ  = items[j].Tag;

            // 入れ替え
            items[i].SubItems[1].Text = nameJ;
            items[i].SubItems[2].Text = typeJ;
            items[i].SubItems[3].Text = zJ;
            items[i].Tag = tagJ;

            items[j].SubItems[1].Text = nameI;
            items[j].SubItems[2].Text = typeI;
            items[j].SubItems[3].Text = zI;
            items[j].Tag = tagI;

            RenumberItems();
        }

        /// <summary>選択行に応じて上下ボタンの有効/無効を切り替える</summary>
        private void UpdateMoveButtons(int selectedIndex)
        {
            btnMoveUp.Enabled   = selectedIndex > 0;
            btnMoveDown.Enabled = selectedIndex >= 0 && selectedIndex < listViewShapes.Items.Count - 1;
        }

        private void ListViewShapes_SelectedIndexChanged(object sender, EventArgs e)
        {
            int idx = listViewShapes.SelectedIndices.Count > 0
                ? listViewShapes.SelectedIndices[0] : -1;
            UpdateMoveButtons(idx);

            if (idx < 0 || idx >= listViewShapes.Items.Count)
            {
                _thumbDebounceTimer.Stop();
                ShowThumbnail(null, null);
                return;
            }

            var shape = listViewShapes.Items[idx].Tag as PowerPoint.Shape;

            // キャッシュヒット → 即時表示（ラグなし）
            if (shape != null && _thumbnailCache.TryGetValue(shape.Name, out var cached))
            {
                _thumbDebounceTimer.Stop();
                ShowThumbnail(cached, shape.Name);
                return;
            }

            // キャッシュミス → 連打中はExportを保留し、止まってから生成
            _pendingThumbIndex = idx;
            _thumbDebounceTimer.Stop();
            _thumbDebounceTimer.Start();
        }

        /// <summary>デバウンスタイマー発火：静止後にサムネイルを生成してキャッシュ</summary>
        private void ThumbDebounceTimer_Tick(object sender, EventArgs e)
        {
            _thumbDebounceTimer.Stop();

            int idx = _pendingThumbIndex;
            if (idx < 0 || idx >= listViewShapes.Items.Count) return;

            var shape = listViewShapes.Items[idx].Tag as PowerPoint.Shape;
            if (shape == null) return;

            // 現在の選択と一致するときだけ生成（遅延中に選択が変わった場合は無視）
            int currentIdx = listViewShapes.SelectedIndices.Count > 0
                ? listViewShapes.SelectedIndices[0] : -1;
            if (currentIdx != idx) return;

            var img = GenerateThumbnail(shape);
            if (img != null) _thumbnailCache[shape.Name] = img;
            ShowThumbnail(img, shape.Name);
        }

        #endregion

        #region OKハンドラ

        private void BtnOK_Click(object sender, EventArgs e)
        {
            // リストの現在順序から OrderedShapes を再構築
            OrderedShapes = new List<PowerPoint.Shape>();
            foreach (ListViewItem item in listViewShapes.Items)
            {
                OrderedShapes.Add((PowerPoint.Shape)item.Tag);
            }

            // 後続アクションを確定（常に選び直し、Z順は変更しない）
            SelectedAction = rbReSelectAndSave.Checked
                ? PostAction.ReSelectAndSaveToStack
                : PostAction.ReSelectOnly;
        }

        #endregion

        #region サムネイル

        /// <summary>PictureBox と名前ラベルを更新する（キャッシュ画像を直接セット）</summary>
        private void ShowThumbnail(Image img, string name)
        {
            pboxThumb.Image = img;
            lblThumbName.Text = img != null ? name : "（未選択）";
        }

        /// <summary>Shape.Export() で図形のサムネイル画像を生成して返す</summary>
        private Image GenerateThumbnail(PowerPoint.Shape shape)
        {
            string tempFile = null;
            try
            {
                tempFile = Path.Combine(Path.GetTempPath(),
                    $"magosa_thumb_{Guid.NewGuid():N}.png");

                shape.Export(tempFile, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

                if (!File.Exists(tempFile)) return null;

                byte[] bytes = File.ReadAllBytes(tempFile);
                using (var ms = new System.IO.MemoryStream(bytes))
                using (var bmp = new Bitmap(ms))
                {
                    // ストリームに依存しない独立したコピーを返す
                    return new Bitmap(bmp);
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (tempFile != null && File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { }
                }
            }
        }

        #endregion

        #region リソース管理

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _thumbDebounceTimer?.Stop();
                _thumbDebounceTimer?.Dispose();
                _thumbDebounceTimer = null;

                // キャッシュ内の全画像を破棄
                pboxThumb.Image = null;
                foreach (var img in _thumbnailCache.Values)
                    img?.Dispose();
                _thumbnailCache.Clear();
            }
            base.Dispose(disposing);
        }

        #endregion

        #region 内部クラス

        /// <summary>
        /// クリックしてもフォーカスを奪わないボタン。
        /// WM_MOUSEACTIVATE に MA_NOACTIVATE を返すことで実現する。
        /// </summary>
        private class NoFocusButton : Button
        {
            protected override void WndProc(ref Message m)
            {
                // WM_MOUSEACTIVATE (0x0021): MA_NOACTIVATE (3) を返してフォーカス遷移を抑止
                if (m.Msg == 0x0021) { m.Result = (IntPtr)3; return; }
                base.WndProc(ref m);
            }
        }

        #endregion
    }
}
