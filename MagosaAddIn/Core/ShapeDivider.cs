using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;

namespace MagosaAddIn.Core
{
    public class ShapeDivider
    {
        public void DivideShape(PowerPoint.Shape originalShape, int rows, int columns,
            float horizontalMargin, float verticalMargin)
        {
            try
            {
                var slide = originalShape.Parent as PowerPoint.Slide;
                if (slide == null)
                {
                    throw new InvalidOperationException("図形が有効なスライドに配置されていません。");
                }

                // 元の図形の位置とサイズを取得
                float originalLeft = originalShape.Left;
                float originalTop = originalShape.Top;
                float originalWidth = originalShape.Width;
                float originalHeight = originalShape.Height;

                // 分割後の各セルのサイズを計算
                float cellWidth = (originalWidth - horizontalMargin * (columns - 1)) / columns;
                float cellHeight = (originalHeight - verticalMargin * (rows - 1)) / rows;

                // サイズが有効かチェック
                if (cellWidth <= 0 || cellHeight <= 0)
                {
                    throw new ArgumentException("マージンが大きすぎるか、分割数が多すぎます。");
                }

                // 元の図形のスタイルを保存
                var shapeStyle = ExtractShapeStyle(originalShape);

                // 分割された図形を作成
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < columns; col++)
                    {
                        float left = originalLeft + col * (cellWidth + horizontalMargin);
                        float top = originalTop + row * (cellHeight + verticalMargin);

                        // 新しい四角形を作成
                        var newShape = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            left, top, cellWidth, cellHeight);

                        // 元の図形のスタイルを適用
                        ApplyShapeStyle(newShape, shapeStyle);

                        // オプション: セル番号をテキストとして追加
                        //if (newShape.TextFrame.HasText == Office.MsoTriState.msoFalse)
                        //{
                        //    newShape.TextFrame.TextRange.Text = $"{row + 1}-{col + 1}";
                        //    var fontSize = Math.Min(cellWidth, cellHeight) / 10;
                        //    if (fontSize > 6) // 最小フォントサイズを設定
                        //    {
                        //        newShape.TextFrame.TextRange.Font.Size = fontSize;
                        //    }
                        //    else
                        //    {
                        //        newShape.TextFrame.TextRange.Font.Size = 6;
                        //    }
                        //}
                    }
                }

                // 元の図形を削除
                originalShape.Delete();
            }
            catch (Exception ex)
            {
                throw new Exception($"図形分割中にエラーが発生しました: {ex.Message}");
            }
        }

        private ShapeStyle ExtractShapeStyle(PowerPoint.Shape shape)
        {
            var style = new ShapeStyle();

            try
            {
                // 塗りつぶし情報
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    style.FillColor = shape.Fill.ForeColor.RGB;
                    style.FillTransparency = shape.Fill.Transparency;
                }

                // 線の情報
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    style.LineColor = shape.Line.ForeColor.RGB;
                    style.LineWeight = shape.Line.Weight;
                    style.LineDashStyle = shape.Line.DashStyle;
                }

                // 影の情報
                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                {
                    style.HasShadow = true;
                    style.ShadowColor = shape.Shadow.ForeColor.RGB;
                }
            }
            catch (Exception ex)
            {
                // スタイル取得でエラーが発生した場合はデフォルト値を使用
                System.Diagnostics.Debug.WriteLine($"スタイル取得エラー: {ex.Message}");
            }

            return style;
        }

        private void ApplyShapeStyle(PowerPoint.Shape shape, ShapeStyle style)
        {
            try
            {
                // 塗りつぶしを適用
                if (style.FillColor.HasValue)
                {
                    shape.Fill.Visible = Office.MsoTriState.msoTrue;
                    shape.Fill.ForeColor.RGB = style.FillColor.Value;
                    shape.Fill.Transparency = style.FillTransparency;
                }

                // 線を適用
                if (style.LineColor.HasValue)
                {
                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = style.LineColor.Value;
                    shape.Line.Weight = style.LineWeight;
                    shape.Line.DashStyle = style.LineDashStyle;
                }

                // 影を適用
                if (style.HasShadow)
                {
                    shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                    shape.Shadow.ForeColor.RGB = style.ShadowColor;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"スタイル適用エラー: {ex.Message}");
            }
        }
    }
}
