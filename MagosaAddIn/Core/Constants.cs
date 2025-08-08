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
    }
}