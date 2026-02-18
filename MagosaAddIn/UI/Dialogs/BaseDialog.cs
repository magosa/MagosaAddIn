using System;
using System.Drawing;
using System.Windows.Forms;

namespace MagosaAddIn.UI.Dialogs
{
    /// <summary>
    /// 全ダイアログの基底クラス
    /// 共通機能とUIコンポーネント作成メソッドを提供
    /// </summary>
    public abstract partial class BaseDialog : Form
    {
        #region 定数定義
        
        // レイアウト定数
        protected new const int DefaultMargin = 20;
        protected const int ControlSpacing = 30;
        protected const int LabelControlGap = 100;
        protected const int ButtonWidth = 75;
        protected const int ButtonHeight = 25;
        protected const int DefaultControlHeight = 20;
        protected const int NumericUpDownWidth = 80;
        
        // ボタン配置用定数
        protected const int ButtonBottomMargin = 60;  // ボタン下部のマージン（ボタン高さを含む総マージン）
        protected const int ButtonTopMargin = 30;      // ボタン上部のマージン（コンテンツとボタンの間）
        protected const int ButtonSpacing = 10;        // ボタン間のスペース
        protected const int ButtonRightMargin = 20;    // ボタン右端のマージン
        
        // 初期Y座標（フォーム上部からの開始位置）
        protected const int InitialTopMargin = 20;
        
        // コントロール間の標準スペース
        protected const int StandardVerticalSpacing = 30;
        protected const int SmallVerticalSpacing = 10;
        
        // フォントスタイル
        protected static readonly Font BoldFont = new Font(SystemFonts.DefaultFont, FontStyle.Bold);
        protected static readonly Font ItalicFont = new Font(SystemFonts.DefaultFont, FontStyle.Italic);
        protected static readonly Font SmallFont = new Font(SystemFonts.DefaultFont.FontFamily, 8);
        
        #endregion
        
        #region 共通プロパティ
        
        protected Button BtnOK { get; set; }
        protected Button BtnCancel { get; set; }
        
        #endregion
        
        #region 共通メソッド - フォーム設定
        
        /// <summary>
        /// フォームの基本設定を行う
        /// </summary>
        /// <param name="title">フォームタイトル</param>
        /// <param name="width">フォーム幅</param>
        /// <param name="height">フォーム高さ</param>
        protected void ConfigureForm(string title, int width, int height)
        {
            this.Text = title;
            this.Size = new Size(width, height);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }
        
        #endregion
        
        #region 共通メソッド - コントロール作成
        
        /// <summary>
        /// ラベルを作成
        /// </summary>
        protected Label CreateLabel(string text, Point location, int width = 80, 
            ContentAlignment textAlign = ContentAlignment.MiddleLeft)
        {
            return new Label
            {
                Text = text,
                Location = location,
                Size = new Size(width, DefaultControlHeight),
                TextAlign = textAlign
            };
        }
        
        /// <summary>
        /// 情報表示用ラベル（太字）を作成
        /// </summary>
        protected Label CreateInfoLabel(string text, Point location, Size size, Color? color = null)
        {
            return new Label
            {
                Text = text,
                Location = location,
                Size = size,
                ForeColor = color ?? Color.DarkBlue,
                Font = BoldFont
            };
        }
        
        /// <summary>
        /// NumericUpDown を作成
        /// </summary>
        protected NumericUpDown CreateNumericUpDown(
            decimal min, decimal max, decimal value,
            Point location, int decimalPlaces = 0, decimal? increment = null)
        {
            return new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Value = value,
                Location = location,
                Size = new Size(NumericUpDownWidth, DefaultControlHeight),
                DecimalPlaces = decimalPlaces,
                Increment = increment ?? (decimalPlaces > 0 ? 0.01m : 1m)
            };
        }
        
        /// <summary>
        /// ボタンを作成
        /// </summary>
        protected Button CreateButton(string text, Point location,
            DialogResult result, EventHandler clickHandler = null)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(ButtonWidth, ButtonHeight),
                DialogResult = result
            };
            if (clickHandler != null)
                btn.Click += clickHandler;
            return btn;
        }
        
        /// <summary>
        /// OK/Cancelボタンを標準配置で追加
        /// </summary>
        /// <param name="yPosition">ボタンのY座標</param>
        /// <param name="okHandler">OKボタンのイベントハンドラ</param>
        /// <param name="buttonWidth">ボタン幅（デフォルト: 75）</param>
        /// <param name="buttonHeight">ボタン高さ（デフォルト: 25）</param>
        protected void AddStandardButtons(int yPosition, EventHandler okHandler, 
            int buttonWidth = ButtonWidth, int buttonHeight = ButtonHeight)
        {
            // ボタン配置の計算：右端から ButtonRightMargin、2つのボタン + 間のスペース
            int xPosition = this.ClientSize.Width - (buttonWidth * 2 + ButtonSpacing + ButtonRightMargin);
            
            BtnOK = new Button
            {
                Text = "OK",
                Location = new Point(xPosition, yPosition),
                Size = new Size(buttonWidth, buttonHeight),
                DialogResult = DialogResult.OK
            };
            BtnOK.Click += okHandler;
            
            BtnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(xPosition + buttonWidth + ButtonSpacing, yPosition),
                Size = new Size(buttonWidth, buttonHeight),
                DialogResult = DialogResult.Cancel
            };
            
            this.Controls.AddRange(new Control[] { BtnOK, BtnCancel });
            this.AcceptButton = BtnOK;
            this.CancelButton = BtnCancel;
        }
        
        /// <summary>
        /// フォーム高さを計算（ボタン位置からフォーム高さを自動計算）
        /// </summary>
        protected int CalculateFormHeight(int buttonY, int buttonHeight = 0)
        {
            int effectiveButtonHeight = buttonHeight > 0 ? buttonHeight : ButtonHeight;
            return buttonY + effectiveButtonHeight + ButtonBottomMargin;
        }
        
        /// <summary>
        /// グループボックスを作成
        /// </summary>
        protected GroupBox CreateGroupBox(string text, Point location, Size size)
        {
            return new GroupBox
            {
                Text = text,
                Location = location,
                Size = size
            };
        }
        
        /// <summary>
        /// チェックボックスを作成
        /// </summary>
        protected CheckBox CreateCheckBox(string text, Point location, Size size, bool isChecked = false)
        {
            return new CheckBox
            {
                Text = text,
                Location = location,
                Size = size,
                Checked = isChecked
            };
        }
        
        /// <summary>
        /// ラジオボタンを作成
        /// </summary>
        protected RadioButton CreateRadioButton(string text, Point location, Size size, bool isChecked = false)
        {
            return new RadioButton
            {
                Text = text,
                Location = location,
                Size = size,
                Checked = isChecked
            };
        }
        
        /// <summary>
        /// コンボボックスを作成
        /// </summary>
        protected ComboBox CreateComboBox(Point location, Size size, ComboBoxStyle style = ComboBoxStyle.DropDownList)
        {
            return new ComboBox
            {
                Location = location,
                Size = size,
                DropDownStyle = style
            };
        }
        
        #endregion
        
        #region リソース管理
        
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                BtnOK?.Dispose();
                BtnCancel?.Dispose();
            }
            base.Dispose(disposing);
        }
        
        #endregion
    }
}
