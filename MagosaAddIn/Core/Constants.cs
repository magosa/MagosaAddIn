namespace MagosaAddIn.Core
{
    /// <summary>
    /// アプリケーション全体で使用する定数を定義するクラス
    /// </summary>
    public static class Constants
    {
        #region 座標・サイズ関連

        /// <summary>
        /// 座標の最小値（pt）
        /// </summary>
        public const float MIN_COORDINATE = -10000.0f;

        /// <summary>
        /// 座標の最大値（pt）
        /// </summary>
        public const float MAX_COORDINATE = 10000.0f;

        /// <summary>
        /// セルの最小サイズ（pt）
        /// </summary>
        public const float MIN_CELL_SIZE = 5.0f;

        /// <summary>
        /// 図形の最小幅（pt）
        /// </summary>
        public const float MIN_SHAPE_WIDTH = 1.0f;

        /// <summary>
        /// 図形の最小高さ（pt）
        /// </summary>
        public const float MIN_SHAPE_HEIGHT = 1.0f;

        #endregion

        #region デフォルト値

        /// <summary>
        /// デフォルト行数
        /// </summary>
        public const int DEFAULT_ROWS = 2;

        /// <summary>
        /// デフォルト列数
        /// </summary>
        public const int DEFAULT_COLUMNS = 2;

        /// <summary>
        /// デフォルト水平マージン（pt）
        /// </summary>
        public const float DEFAULT_HORIZONTAL_MARGIN = 10.0f;

        /// <summary>
        /// デフォルト垂直マージン（pt）
        /// </summary>
        public const float DEFAULT_VERTICAL_MARGIN = 10.0f;

        /// <summary>
        /// デフォルト線の太さ（pt）
        /// </summary>
        public const float DEFAULT_LINE_WEIGHT = 1.0f;

        /// <summary>
        /// デフォルト半径（pt）
        /// </summary>
        public const float DEFAULT_RADIUS = 100.0f;

        /// <summary>
        /// デフォルト中心X座標（pt）
        /// </summary>
        public const float DEFAULT_CENTER_X = 400.0f;

        /// <summary>
        /// デフォルト中心Y座標（pt）
        /// </summary>
        public const float DEFAULT_CENTER_Y = 300.0f;

        #endregion

        #region 制限値

        /// <summary>
        /// 行数の最小値
        /// </summary>
        public const int MIN_ROWS = 1;

        /// <summary>
        /// 行数の最大値
        /// </summary>
        public const int MAX_ROWS = 50;

        /// <summary>
        /// 列数の最小値
        /// </summary>
        public const int MIN_COLUMNS = 1;

        /// <summary>
        /// 列数の最大値
        /// </summary>
        public const int MAX_COLUMNS = 50;

        /// <summary>
        /// マージンの最小値（pt）
        /// </summary>
        public const float MIN_MARGIN = 0.0f;

        /// <summary>
        /// マージンの最大値（pt）
        /// </summary>
        public const float MAX_MARGIN = 200.0f;

        /// <summary>
        /// 半径の最小値（pt）
        /// </summary>
        public const float MIN_RADIUS = 10.0f;

        /// <summary>
        /// 半径の最大値（pt）
        /// </summary>
        public const float MAX_RADIUS = 500.0f;

        /// <summary>
        /// 中心座標の最小値（pt）
        /// </summary>
        public const float MIN_CENTER_COORDINATE = -1000.0f;

        /// <summary>
        /// 中心座標の最大値（pt）
        /// </summary>
        public const float MAX_CENTER_COORDINATE = 2000.0f;

        /// <summary>
        /// 間隔の最大値（pt）
        /// </summary>
        public const float MAX_SPACING = 200.0f;

        #endregion

        #region 色関連

        /// <summary>
        /// デフォルト塗りつぶし色（白）
        /// </summary>
        public const int DEFAULT_FILL_COLOR = 0xFFFFFF;

        /// <summary>
        /// デフォルト線色（黒）
        /// </summary>
        public const int DEFAULT_LINE_COLOR = 0x000000;

        /// <summary>
        /// デフォルト透明度
        /// </summary>
        public const float DEFAULT_TRANSPARENCY = 0.0f;

        #endregion

        #region 図形選択関連

        /// <summary>
        /// 整列に必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_ALIGNMENT = 2;

        /// <summary>
        /// 分割に必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_DIVISION = 1;

        #endregion

        #region UI関連

        /// <summary>
        /// NumericUpDownの小数点桁数
        /// </summary>
        public const int DECIMAL_PLACES = 1;

        /// <summary>
        /// NumericUpDownの増減値
        /// </summary>
        public const decimal INCREMENT_VALUE = 0.5m;

        /// <summary>
        /// デフォルトグリッド列数
        /// </summary>
        public const int DEFAULT_GRID_COLUMNS = 3;

        #endregion

        #region 選択補助関連

        /// <summary>
        /// 書式選択に必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_FORMAT_SELECTION = 1;

        #endregion

        #region 図形置き換え関連

        /// <summary>
        /// 図形置き換えに必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_REPLACEMENT = 1;

        /// <summary>
        /// テンプレート図形選択に必要な図形数
        /// </summary>
        public const int TEMPLATE_SHAPE_COUNT = 1;

        #endregion

        #region 図形ハンドル調整関連

        /// <summary>
        /// デフォルト調整ハンドル値（正規化値）
        /// </summary>
        public const float DEFAULT_HANDLE_VALUE = 0.5f;

        /// <summary>
        /// 調整ハンドル値の最小値（正規化値）
        /// </summary>
        public const float MIN_HANDLE_VALUE = 0.0f;

        /// <summary>
        /// 調整ハンドル値の最大値（正規化値）
        /// </summary>
        public const float MAX_HANDLE_VALUE = 1.0f;

        /// <summary>
        /// ハンドル調整に必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_HANDLE_ADJUSTMENT = 1;

        /// <summary>
        /// サポートする最大調整ハンドル数
        /// </summary>
        public const int MAX_SUPPORTED_HANDLES = 8;

        /// <summary>
        /// デフォルト調整値（mm）
        /// </summary>
        public const float DEFAULT_HANDLE_MM = 3.0f;

        /// <summary>
        /// 調整値の最小値（mm）
        /// </summary>
        public const float MIN_HANDLE_MM = 0.0f;

        /// <summary>
        /// 調整値の最大値（mm）
        /// </summary>
        public const float MAX_HANDLE_MM = 50.0f;

        /// <summary>
        /// デフォルト角度値（度）
        /// </summary>
        public const float DEFAULT_ANGLE_DEGREE = 24.0f;

        /// <summary>
        /// 角度の最小値（度）
        /// </summary>
        public const float MIN_ANGLE_DEGREE = 0.0f;

        /// <summary>
        /// 角度の最大値（度）
        /// </summary>
        public const float MAX_ANGLE_DEGREE = 360.0f;

        /// <summary>
        /// mm→pt変換係数（1mm = 2.834645669pt）
        /// </summary>
        public const float MM_TO_PT = 2.834645669f;

        /// <summary>
        /// pt→mm変換係数
        /// </summary>
        public const float PT_TO_MM = 1.0f / MM_TO_PT;

        #endregion

        #region レイヤー調整関連

        /// <summary>
        /// レイヤー調整に必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_LAYER = 2;

        #endregion

        #region 自動ナンバリング関連

        /// <summary>
        /// 自動ナンバリングに必要な最小図形数
        /// </summary>
        public const int MIN_SHAPES_FOR_NUMBERING = 1;

        /// <summary>
        /// デフォルト開始番号
        /// </summary>
        public const int DEFAULT_START_NUMBER = 1;

        /// <summary>
        /// デフォルト増分値
        /// </summary>
        public const int DEFAULT_INCREMENT = 1;

        /// <summary>
        /// 開始番号の最小値
        /// </summary>
        public const int MIN_START_NUMBER = 0;

        /// <summary>
        /// 開始番号の最大値
        /// </summary>
        public const int MAX_START_NUMBER = 999;

        /// <summary>
        /// 増分値の最小値
        /// </summary>
        public const int MIN_INCREMENT = -10;

        /// <summary>
        /// 増分値の最大値
        /// </summary>
        public const int MAX_INCREMENT = 10;

        /// <summary>
        /// デフォルトフォントサイズ（pt）
        /// </summary>
        public const float DEFAULT_NUMBER_FONT_SIZE = 18.0f;

        #endregion
    }
}
