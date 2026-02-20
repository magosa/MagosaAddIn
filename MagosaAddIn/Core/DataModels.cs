using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// レイヤー（重なり順）の調整方向
    /// </summary>
    public enum LayerOrder
    {
        /// <summary>選択順に前面へ配置</summary>
        SelectionOrderToFront,
        /// <summary>選択順に背面へ配置</summary>
        SelectionOrderToBack,
        /// <summary>左から右へ前面に配置</summary>
        LeftToRightToFront,
        /// <summary>上から下へ前面に配置</summary>
        TopToBottomToFront
    }

    /// <summary>
    /// 自動ナンバリングのフォーマット形式
    /// </summary>
    public enum NumberFormat
    {
        /// <summary>算用数字（1, 2, 3...）</summary>
        Arabic,
        /// <summary>丸数字（①②③...）</summary>
        CircledArabic,
        /// <summary>大文字アルファベット（A, B, C...）</summary>
        UpperAlpha,
        /// <summary>小文字アルファベット（a, b, c...）</summary>
        LowerAlpha,
        /// <summary>ローマ数字大文字（I, II, III...）</summary>
        UpperRoman,
        /// <summary>ローマ数字小文字（i, ii, iii...）</summary>
        LowerRoman
    }

    /// <summary>
    /// サイズの単位
    /// </summary>
    public enum SizeUnit
    {
        /// <summary>ポイント（pt）</summary>
        Point,
        /// <summary>ミリメートル（mm）</summary>
        Millimeter,
        /// <summary>センチメートル（cm）</summary>
        Centimeter
    }

    /// <summary>
    /// リサイズモード
    /// </summary>
    public enum ResizeMode
    {
        /// <summary>中心位置保持</summary>
        KeepCenter,
        /// <summary>左上位置保持</summary>
        KeepTopLeft
    }

    /// <summary>
    /// 図形情報を格納するクラス
    /// </summary>
    public class ShapeInfo
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string Name { get; set; }
        public string ShapeName { get; set; }
        public PowerPoint.Shape OriginalShape { get; set; }

        // 中心座標
        public float CenterX { get; set; }
        public float CenterY { get; set; }

        // スタイル情報
        public int? FillColor { get; set; }
        public float FillTransparency { get; set; }
        public int? LineColor { get; set; }
        public float LineWeight { get; set; } = Constants.DEFAULT_LINE_WEIGHT;
        public Office.MsoLineDashStyle LineDashStyle { get; set; }
        public bool HasShadow { get; set; }
        public int ShadowColor { get; set; }

        // テキスト情報
        public string Text { get; set; }

        /// <summary>
        /// 図形情報の文字列表現
        /// </summary>
        public override string ToString()
        {
            return $"{Name}: ({Left:F1}, {Top:F1}) {Width:F1}×{Height:F1}";
        }

        /// <summary>
        /// 図形の中心座標を取得
        /// </summary>
        public (float X, float Y) GetCenter()
        {
            return (Left + Width / 2, Top + Height / 2);
        }

        /// <summary>
        /// 図形の右端座標を取得
        /// </summary>
        public float GetRight()
        {
            return Left + Width;
        }

        /// <summary>
        /// 図形の下端座標を取得
        /// </summary>
        public float GetBottom()
        {
            return Top + Height;
        }
    }

    /// <summary>
    /// 図形グループの境界情報
    /// </summary>
    public class ShapeGroupBounds
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Right { get; set; }
        public float Bottom { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }

        /// <summary>
        /// 境界情報の文字列表現
        /// </summary>
        public override string ToString()
        {
            return $"Bounds: ({Left:F1}, {Top:F1}) - ({Right:F1}, {Bottom:F1}) Size: {Width:F1}×{Height:F1}";
        }

        /// <summary>
        /// 境界の中心座標を取得
        /// </summary>
        public (float X, float Y) GetCenter()
        {
            return ((Left + Right) / 2, (Top + Bottom) / 2);
        }

        /// <summary>
        /// 境界の面積を取得
        /// </summary>
        public float GetArea()
        {
            return Width * Height;
        }

        /// <summary>
        /// 指定した座標が境界内にあるかチェック
        /// </summary>
        public bool Contains(float x, float y)
        {
            return x >= Left && x <= Right && y >= Top && y <= Bottom;
        }

        /// <summary>
        /// 図形リストから境界を計算
        /// </summary>
        public static ShapeGroupBounds FromShapes(System.Collections.Generic.List<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count == 0)
                throw new System.ArgumentException("図形リストが空です。");

            float minLeft = shapes.Min(s => s.Left);
            float minTop = shapes.Min(s => s.Top);
            float maxRight = shapes.Max(s => s.Left + s.Width);
            float maxBottom = shapes.Max(s => s.Top + s.Height);

            return new ShapeGroupBounds
            {
                Left = minLeft,
                Top = minTop,
                Right = maxRight,
                Bottom = maxBottom,
                Width = maxRight - minLeft,
                Height = maxBottom - minTop
            };
        }

        /// <summary>
        /// ShapeInfoリストから境界を計算
        /// </summary>
        public static ShapeGroupBounds FromShapeInfos(System.Collections.Generic.List<ShapeInfo> shapeInfos)
        {
            if (shapeInfos == null || shapeInfos.Count == 0)
                throw new System.ArgumentException("図形情報リストが空です。");

            float minLeft = shapeInfos.Min(s => s.Left);
            float minTop = shapeInfos.Min(s => s.Top);
            float maxRight = shapeInfos.Max(s => s.GetRight());
            float maxBottom = shapeInfos.Max(s => s.GetBottom());

            return new ShapeGroupBounds
            {
                Left = minLeft,
                Top = minTop,
                Right = maxRight,
                Bottom = maxBottom,
                Width = maxRight - minLeft,
                Height = maxBottom - minTop
            };
        }
    }

    #region 配列複製オプションクラス群

    /// <summary>
    /// 線形配列のオプション
    /// </summary>
    public class LinearArrayOptions
    {
        /// <summary>配列方向（度）</summary>
        public float Angle { get; set; }
        /// <summary>個数</summary>
        public int Count { get; set; }
        /// <summary>間隔（pt）</summary>
        public float Spacing { get; set; }
    }

    /// <summary>
    /// 配列の角度モード
    /// </summary>
    public enum ArrayAngleMode
    {
        /// <summary>等分配置（360° ÷ 個数）</summary>
        EqualDivision,
        /// <summary>角度指定（指定角度 × インデックス）</summary>
        AngleStep
    }

    /// <summary>
    /// 回転中心の指定方法
    /// </summary>
    public enum CenterSource
    {
        /// <summary>カスタム座標</summary>
        CustomCoordinate,
        /// <summary>配列対象図形の中心</summary>
        TargetShapeCenter
    }

    /// <summary>
    /// 円形配列のオプション（回転コピー統合版）
    /// </summary>
    public class CircularArrayOptions
    {
        /// <summary>配列モード</summary>
        public ArrayAngleMode AngleMode { get; set; }
        
        /// <summary>回転中心の指定方法</summary>
        public CenterSource CenterSource { get; set; }
        
        /// <summary>中心X座標（カスタム座標または選択図形から取得）</summary>
        public float CenterX { get; set; }
        /// <summary>中心Y座標（カスタム座標または選択図形から取得）</summary>
        public float CenterY { get; set; }
        
        /// <summary>半径（等分配置モード用）</summary>
        public float Radius { get; set; }
        
        /// <summary>個数</summary>
        public int Count { get; set; }
        
        /// <summary>開始角度</summary>
        public float StartAngle { get; set; }
        
        /// <summary>角度ステップ（角度指定モード用）</summary>
        public float AngleStep { get; set; }
        
        /// <summary>図形を回転させるか</summary>
        public bool RotateShapes { get; set; }
        
        /// <summary>中心座標取得用の図形（内部使用）</summary>
        public PowerPoint.Shape CenterReferenceShape { get; set; }
    }

    /// <summary>
    /// グリッド配列のオプション（角度対応・線形配列統合版）
    /// </summary>
    public class GridArrayOptions
    {
        /// <summary>行数（1=横方向線形配列）</summary>
        public int Rows { get; set; }
        /// <summary>列数（1=縦方向線形配列）</summary>
        public int Columns { get; set; }
        /// <summary>水平間隔（pt）</summary>
        public float HorizontalSpacing { get; set; }
        /// <summary>垂直間隔（pt）</summary>
        public float VerticalSpacing { get; set; }
        /// <summary>グリッド全体の回転角度（度）0°=右, 90°=下</summary>
        public float Angle { get; set; }
    }

    /// <summary>
    /// パス配列のオプション
    /// </summary>
    public class PathArrayOptions
    {
        /// <summary>個数</summary>
        public int Count { get; set; }
        /// <summary>等間隔配置</summary>
        public bool EqualSpacing { get; set; }
        /// <summary>カスタム間隔</summary>
        public float CustomSpacing { get; set; }
        /// <summary>パスに沿って回転</summary>
        public bool RotateAlongPath { get; set; }
    }

    /// <summary>
    /// 回転コピーのオプション
    /// </summary>
    public class RotationCopyOptions
    {
        /// <summary>回転中心X</summary>
        public float CenterX { get; set; }
        /// <summary>回転中心Y</summary>
        public float CenterY { get; set; }
        /// <summary>回転角度（度）</summary>
        public float Angle { get; set; }
        /// <summary>個数</summary>
        public int Count { get; set; }
        /// <summary>図形中心を使用</summary>
        public bool UseShapeCenter { get; set; }
    }

    #endregion

    #region テーマカラー生成関連

    /// <summary>
    /// 配色パターンの種類
    /// </summary>
    public enum ColorSchemeType
    {
        // 色相ベース配色
        /// <summary>ダイアード（補色）- 色相環で180°反対</summary>
        Dyad,
        /// <summary>トライアド（3等分）- 色相環で120°間隔</summary>
        Triad,
        /// <summary>テトラード（4等分）- 色相環で90°間隔</summary>
        Tetrad,
        /// <summary>ペンタード（5等分）- 色相環で72°間隔</summary>
        Pentad,
        /// <summary>ヘクサード（6等分）- 色相環で60°間隔</summary>
        Hexad,
        /// <summary>アナロジー（類似色）- 色相環で隣接</summary>
        Analogy,
        /// <summary>インターミディエート - 色相環で90°</summary>
        Intermediate,
        /// <summary>オポーネント - 色相環で135°</summary>
        Opponent,
        /// <summary>スプリットコンプリメンタリー - 補色の分割</summary>
        SplitComplementary,

        // トーンベース配色
        /// <summary>トーンオントーン - 同色相・明度差大</summary>
        ToneOnTone,
        /// <summary>トーンイントーン - トーン統一・色相変化</summary>
        ToneInTone,
        /// <summary>カマイユ - 色相・トーン近似</summary>
        Camaieu,
        /// <summary>フォカマイユ - カマイユより色相変化</summary>
        FauxCamaieu,
        /// <summary>ドミナントカラー - 同色相統一</summary>
        DominantColor,
        /// <summary>アイデンティティ - 1色相・明度彩度変化</summary>
        Identity,
        /// <summary>グラデーション - 段階的変化</summary>
        Gradation,

        // コントラスト配色
        /// <summary>色相コントラスト - 補色関係</summary>
        HueContrast,
        /// <summary>明度コントラスト - 明暗対比</summary>
        LightnessContrast,
        /// <summary>彩度コントラスト - 鮮やか⇔くすみ</summary>
        SaturationContrast
    }

    /// <summary>
    /// テーマカラー生成オプション
    /// </summary>
    public class ThemeColorOptions
    {
        /// <summary>ベースカラー（PowerPoint RGB値）</summary>
        public int BaseColor { get; set; }

        /// <summary>配色パターン</summary>
        public ColorSchemeType SchemeType { get; set; }

        /// <summary>生成する色数（3～10色）</summary>
        public int ColorCount { get; set; } = 5;

        /// <summary>明度バリエーション段階数（2～5段階）</summary>
        public int LightnessSteps { get; set; } = 3;

        /// <summary>パレット配置を行うか</summary>
        public bool ArrangePalette { get; set; } = false;

        /// <summary>選択図形に適用するか</summary>
        public bool ApplyToShapes { get; set; } = false;
    }

    /// <summary>
    /// カラーパレット配置位置
    /// </summary>
    public enum PalettePosition
    {
        /// <summary>スライド右側</summary>
        Right,
        /// <summary>スライド下側</summary>
        Bottom
    }

    /// <summary>
    /// カラーパレット配置オプション
    /// </summary>
    public class PaletteArrangementOptions
    {
        /// <summary>セルサイズ（pt）</summary>
        public float CellSize { get; set; } = 20f;

        /// <summary>配置位置</summary>
        public PalettePosition Position { get; set; } = PalettePosition.Right;

        /// <summary>スライドとの間隔（pt）</summary>
        public float Margin { get; set; } = 20f;
    }

    #endregion
}
