namespace MagosaAddIn.Core
{
    /// <summary>
    /// 画像色編集の設定値を保持する DTO クラス
    /// </summary>
    public class ImageColorSettings
    {
        /// <summary>明るさ補正値（-100〜+100、0=変更なし）</summary>
        public int Brightness { get; set; } = 0;

        /// <summary>コントラスト補正値（-100〜+100、0=変更なし）</summary>
        public int Contrast { get; set; } = 0;

        /// <summary>色相シフト値（-180〜+180、0=変更なし）</summary>
        public int Hue { get; set; } = 0;

        /// <summary>彩度補正値（-100〜+100、0=変更なし）</summary>
        public int Saturation { get; set; } = 0;

        /// <summary>カラーライズを有効にするか</summary>
        public bool ColorizeEnabled { get; set; } = false;

        /// <summary>カラーライズ強度（0〜100）</summary>
        public int ColorizeIntensity { get; set; } = 50;

        /// <summary>カラーライズ色（PowerPoint RGB 形式 0xBBGGRR、R=低バイト）</summary>
        public int ColorizeRgb { get; set; } = 0xFF; // デフォルト: 赤 (R=255, G=0, B=0)

        /// <summary>色調変換モード</summary>
        public ColorToneMode ToneMode { get; set; } = ColorToneMode.None;

        /// <summary>白黒変換の閾値（0〜100）</summary>
        public int BlackWhiteThreshold { get; set; } = 50;
    }

    /// <summary>色調変換モード</summary>
    public enum ColorToneMode
    {
        /// <summary>変換なし</summary>
        None,
        /// <summary>グレースケール</summary>
        Grayscale,
        /// <summary>セピア</summary>
        Sepia,
        /// <summary>白黒（二値化）</summary>
        BlackAndWhite
    }
}
