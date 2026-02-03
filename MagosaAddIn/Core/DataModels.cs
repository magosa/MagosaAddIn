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
}