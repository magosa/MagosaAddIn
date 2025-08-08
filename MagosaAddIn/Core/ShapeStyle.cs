using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形スタイルを保存するためのクラス
    /// </summary>
    public class ShapeStyle
    {
        public int? FillColor { get; set; }
        public float FillTransparency { get; set; }
        public int? LineColor { get; set; }
        public float LineWeight { get; set; } = Constants.DEFAULT_LINE_WEIGHT;
        public Office.MsoLineDashStyle LineDashStyle { get; set; } = Office.MsoLineDashStyle.msoLineSolid;
        public bool HasShadow { get; set; }
        public int ShadowColor { get; set; }

        /// <summary>
        /// スタイル情報の文字列表現
        /// </summary>
        public override string ToString()
        {
            return $"Fill: {FillColor?.ToString("X6") ?? "None"}, " +
                   $"Line: {LineColor?.ToString("X6") ?? "None"}, " +
                   $"Weight: {LineWeight}, " +
                   $"Shadow: {(HasShadow ? "Yes" : "No")}";
        }

        /// <summary>
        /// デフォルトスタイルを作成
        /// </summary>
        public static ShapeStyle CreateDefault()
        {
            return new ShapeStyle
            {
                FillColor = Constants.DEFAULT_FILL_COLOR,
                FillTransparency = Constants.DEFAULT_TRANSPARENCY,
                LineColor = Constants.DEFAULT_LINE_COLOR,
                LineWeight = Constants.DEFAULT_LINE_WEIGHT,
                LineDashStyle = Office.MsoLineDashStyle.msoLineSolid,
                HasShadow = false
            };
        }
    }
}