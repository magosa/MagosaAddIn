using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagosaAddIn
{
    /// <summary>
    /// 図形の分割機能を提供するクラス
    /// </summary>
    public class ShapeDivider
    {
        #region 単一図形分割機能（既存機能）

        /// <summary>
        /// 単一図形を分割する
        /// </summary>
        /// <param name="originalShape">分割する図形</param>
        /// <param name="rows">行数</param>
        /// <param name="columns">列数</param>
        /// <param name="horizontalMargin">水平マージン</param>
        /// <param name="verticalMargin">垂直マージン</param>
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

        #endregion

        #region 複数図形グリッド分割機能（新機能）

        /// <summary>
        /// 複数図形の範囲内でグリッド分割を実行（情報事前取得方式）
        /// </summary>
        /// <param name="originalShapes">元の図形リスト</param>
        /// <param name="rows">行数</param>
        /// <param name="columns">列数</param>
        /// <param name="horizontalMargin">水平マージン</param>
        /// <param name="verticalMargin">垂直マージン</param>
        /// <param name="deleteOriginalShapes">元図形を削除するかどうか</param>
        public void DivideShapeGroup(List<PowerPoint.Shape> originalShapes, int rows, int columns,
            float horizontalMargin, float verticalMargin, bool deleteOriginalShapes = false)
        {
            try
            {
                if (originalShapes == null || originalShapes.Count == 0)
                {
                    throw new ArgumentException("図形が指定されていません。");
                }

                if (rows <= 0 || columns <= 0)
                {
                    throw new ArgumentException("行数と列数は1以上である必要があります。");
                }

                // 1. 事前に図形情報を取得（COM参照を回避）
                var shapeInfos = ExtractShapeInfos(originalShapes);
                if (shapeInfos.Count == 0)
                {
                    throw new InvalidOperationException("有効な図形情報を取得できませんでした。");
                }

                // 2. スライド参照を取得
                PowerPoint.Slide slide = GetValidSlide(originalShapes);
                if (slide == null)
                {
                    throw new InvalidOperationException("有効なスライドが見つかりません。");
                }

                // 3. 図形グループの境界を計算（事前取得した情報を使用）
                var bounds = CalculateBounds(shapeInfos);

                // 4. 分割後の各セルのサイズを計算
                float cellWidth = (bounds.Width - horizontalMargin * (columns - 1)) / columns;
                float cellHeight = (bounds.Height - verticalMargin * (rows - 1)) / rows;

                // 5. サイズが有効かチェック
                if (cellWidth <= 5 || cellHeight <= 5)
                {
                    throw new ArgumentException($"セルサイズが小さすぎます。計算されたサイズ: {cellWidth:F1}×{cellHeight:F1}pt\n" +
                        $"範囲: {bounds.Width:F1}×{bounds.Height:F1}pt\n" +
                        $"マージンを小さくするか、分割数を減らしてください。");
                }

                // 6. 代表的なスタイルを取得（最初の図形から）
                var shapeStyle = ExtractShapeStyleFromInfo(shapeInfos[0]);

                // 7. グリッド分割を実行
                var createdShapes = CreateGridShapes(slide, bounds, rows, columns,
                    cellWidth, cellHeight, horizontalMargin, verticalMargin, shapeStyle);

                // 8. 元図形を削除（指定された場合）
                if (deleteOriginalShapes)
                {
                    DeleteOriginalShapes(originalShapes);
                }

                System.Diagnostics.Debug.WriteLine($"グリッド分割完了: {rows}×{columns}, " +
                    $"範囲: {bounds.Width:F2}×{bounds.Height:F2}, 作成図形数: {createdShapes.Count}");
            }
            catch (Exception ex)
            {
                throw new Exception($"グリッド分割中にエラーが発生しました: {ex.Message}");
            }
        }

        /// <summary>
        /// 図形情報を事前に取得
        /// </summary>
        private List<ShapeInfo> ExtractShapeInfos(List<PowerPoint.Shape> shapes)
        {
            var shapeInfos = new List<ShapeInfo>();

            foreach (var shape in shapes)
            {
                try
                {
                    var info = new ShapeInfo
                    {
                        Left = shape.Left,
                        Top = shape.Top,
                        Width = shape.Width,
                        Height = shape.Height,
                        Name = shape.Name,
                        OriginalShape = shape
                    };

                    // スタイル情報も事前取得
                    try
                    {
                        info.FillColor = shape.Fill.Visible == Office.MsoTriState.msoTrue ?
                            (int?)shape.Fill.ForeColor.RGB : null;
                        info.FillTransparency = shape.Fill.Transparency;
                        info.LineColor = shape.Line.Visible == Office.MsoTriState.msoTrue ?
                            (int?)shape.Line.ForeColor.RGB : null;
                        info.LineWeight = shape.Line.Weight;
                        info.LineDashStyle = shape.Line.DashStyle;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // スタイル情報の取得に失敗した場合はデフォルト値を使用
                        info.FillColor = 0xFFFFFF; // 白
                        info.LineColor = 0x000000; // 黒
                        info.LineWeight = 1.0f;
                        info.LineDashStyle = Office.MsoLineDashStyle.msoLineSolid;
                    }

                    shapeInfos.Add(info);
                    System.Diagnostics.Debug.WriteLine($"図形情報取得成功: {info.Name} ({info.Left:F1}, {info.Top:F1}, {info.Width:F1}×{info.Height:F1})");
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    System.Diagnostics.Debug.WriteLine($"図形情報取得失敗: {comEx.Message}");
                    continue;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"図形情報取得エラー: {ex.Message}");
                    continue;
                }
            }

            System.Diagnostics.Debug.WriteLine($"取得した図形情報数: {shapeInfos.Count}/{shapes.Count}");
            return shapeInfos;
        }

        /// <summary>
        /// 有効なスライドを取得
        /// </summary>
        private PowerPoint.Slide GetValidSlide(List<PowerPoint.Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                try
                {
                    var slide = shape.Parent as PowerPoint.Slide;
                    if (slide != null)
                    {
                        // スライドの有効性をテスト
                        var testName = slide.Name;
                        return slide;
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    continue;
                }
            }
            return null;
        }

        /// <summary>
        /// 境界を計算（事前取得した情報を使用）
        /// </summary>
        private ShapeGroupBounds CalculateBounds(List<ShapeInfo> shapeInfos)
        {
            if (shapeInfos.Count == 0)
                throw new ArgumentException("図形情報が空です。");

            float minLeft = shapeInfos.Min(s => s.Left);
            float minTop = shapeInfos.Min(s => s.Top);
            float maxRight = shapeInfos.Max(s => s.Left + s.Width);
            float maxBottom = shapeInfos.Max(s => s.Top + s.Height);

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
        /// 事前取得した情報からスタイルを作成
        /// </summary>
        private ShapeStyle ExtractShapeStyleFromInfo(ShapeInfo shapeInfo)
        {
            return new ShapeStyle
            {
                FillColor = shapeInfo.FillColor,
                FillTransparency = shapeInfo.FillTransparency,
                LineColor = shapeInfo.LineColor,
                LineWeight = shapeInfo.LineWeight,
                LineDashStyle = shapeInfo.LineDashStyle,
                HasShadow = false // 影は複雑なのでとりあえず無効
            };
        }

        /// <summary>
        /// グリッド図形を作成
        /// </summary>
        private List<PowerPoint.Shape> CreateGridShapes(PowerPoint.Slide slide, ShapeGroupBounds bounds,
            int rows, int columns, float cellWidth, float cellHeight,
            float horizontalMargin, float verticalMargin, ShapeStyle style)
        {
            var createdShapes = new List<PowerPoint.Shape>();

            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < columns; col++)
                {
                    try
                    {
                        float left = bounds.Left + col * (cellWidth + horizontalMargin);
                        float top = bounds.Top + row * (cellHeight + verticalMargin);

                        // 座標の有効性をチェック
                        if (left < -10000 || left > 10000 || top < -10000 || top > 10000)
                        {
                            System.Diagnostics.Debug.WriteLine($"座標が範囲外: ({left:F1}, {top:F1}) - スキップ");
                            continue;
                        }

                        // 新しい四角形を作成
                        var newShape = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            left, top, cellWidth, cellHeight);

                        createdShapes.Add(newShape);

                        // スタイルを適用
                        ApplyShapeStyleSafe(newShape, style);

                        System.Diagnostics.Debug.WriteLine($"図形作成成功: 行{row + 1}列{col + 1} ({left:F1}, {top:F1})");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"図形作成エラー (行{row + 1}, 列{col + 1}): {ex.Message}");
                        // エラーが発生しても処理を継続
                        continue;
                    }
                }
            }

            return createdShapes;
        }

        /// <summary>
        /// 元図形を安全に削除
        /// </summary>
        private void DeleteOriginalShapes(List<PowerPoint.Shape> shapes)
        {
            int deletedCount = 0;
            foreach (var shape in shapes)
            {
                try
                {
                    shape.Delete();
                    deletedCount++;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    System.Diagnostics.Debug.WriteLine("図形削除スキップ（既に削除済みまたは無効）");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"図形削除エラー: {ex.Message}");
                }
            }
            System.Diagnostics.Debug.WriteLine($"削除した図形数: {deletedCount}/{shapes.Count}");
        }

        #endregion

        #region 共通スタイル処理メソッド

        /// <summary>
        /// 図形のスタイルを抽出する
        /// </summary>
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
                System.Diagnostics.Debug.WriteLine($"スタイル取得エラー: {ex.Message}");
            }

            return style;
        }

        /// <summary>
        /// 図形にスタイルを適用する
        /// </summary>
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

        /// <summary>
        /// 安全な図形スタイル適用
        /// </summary>
        private void ApplyShapeStyleSafe(PowerPoint.Shape shape, ShapeStyle style)
        {
            try
            {
                // 塗りつぶしを適用
                if (style.FillColor.HasValue)
                {
                    try
                    {
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.ForeColor.RGB = style.FillColor.Value;
                        shape.Fill.Transparency = style.FillTransparency;
                    }
                    catch (System.Runtime.InteropServices.COMException) { }
                }

                // 線を適用
                if (style.LineColor.HasValue)
                {
                    try
                    {
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.ForeColor.RGB = style.LineColor.Value;
                        shape.Line.Weight = style.LineWeight;
                        shape.Line.DashStyle = style.LineDashStyle;
                    }
                    catch (System.Runtime.InteropServices.COMException) { }
                }

                // 影を適用
                if (style.HasShadow)
                {
                    try
                    {
                        shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                        shape.Shadow.ForeColor.RGB = style.ShadowColor;
                    }
                    catch (System.Runtime.InteropServices.COMException) { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"スタイル適用エラー: {ex.Message}");
            }
        }

        #endregion
    }

    #region データクラス

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
        public PowerPoint.Shape OriginalShape { get; set; }

        // スタイル情報
        public int? FillColor { get; set; }
        public float FillTransparency { get; set; }
        public int? LineColor { get; set; }
        public float LineWeight { get; set; }
        public Office.MsoLineDashStyle LineDashStyle { get; set; }

        public override string ToString()
        {
            return $"{Name}: ({Left:F1}, {Top:F1}) {Width:F1}×{Height:F1}";
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

        public override string ToString()
        {
            return $"Bounds: ({Left:F1}, {Top:F1}) - ({Right:F1}, {Bottom:F1}) Size: {Width:F1}×{Height:F1}";
        }
    }

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

        public override string ToString()
        {
            return $"Fill: {FillColor?.ToString("X6") ?? "None"}, Line: {LineColor?.ToString("X6") ?? "None"}, Weight: {LineWeight}";
        }
    }

    #endregion
}