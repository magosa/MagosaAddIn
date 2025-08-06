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
        public float LineWeight { get; set; } = 1.0f;
        public Office.MsoLineDashStyle LineDashStyle { get; set; } = Office.MsoLineDashStyle.msoLineSolid;
        public bool HasShadow { get; set; }
        public int ShadowColor { get; set; }
    }
}
