using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using MagosaAddIn.Core;

namespace MagosaAddIn.Core
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
                // 入力検証
                if (originalShape == null)
                {
                    ErrorHandler.ShowOperationError("図形分割", new ArgumentNullException(nameof(originalShape), "分割する図形が指定されていません。"));
                    return;
                }

                ErrorHandler.ValidateRange(rows, Constants.MIN_ROWS, Constants.MAX_ROWS, "行数", "図形分割");
                ErrorHandler.ValidateRange(columns, Constants.MIN_COLUMNS, Constants.MAX_COLUMNS, "列数", "図形分割");
                ErrorHandler.ValidateRange(horizontalMargin, Constants.MIN_MARGIN, Constants.MAX_MARGIN, "水平マージン", "図形分割");
                ErrorHandler.ValidateRange(verticalMargin, Constants.MIN_MARGIN, Constants.MAX_MARGIN, "垂直マージン", "図形分割");

                var slide = ComExceptionHandler.ExecuteComOperation(
                    () => originalShape.Parent as PowerPoint.Slide,
                    "スライド取得");

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

                // セルサイズの検証
                ErrorHandler.ValidateCellSize(cellWidth, cellHeight, "図形分割");

                // 元の図形のスタイルを保存
                var shapeStyle = ExtractShapeStyle(originalShape);

                // 分割された図形を作成
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < columns; col++)
                    {
                        float left = originalLeft + col * (cellWidth + horizontalMargin);
                        float top = originalTop + row * (cellHeight + verticalMargin);

                        // 座標の検証
                        if (!ErrorHandler.ValidateCoordinates(left, top, "図形分割"))
                        {
                            continue;
                        }

                        // 新しい四角形を作成
                        var newShape = ComExceptionHandler.ExecuteComOperation(
                            () => slide.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeRectangle,
                                left, top, cellWidth, cellHeight),
                            $"図形作成 (行{row + 1}, 列{col + 1})");

                        // 元の図形のスタイルを適用
                        ApplyShapeStyle(newShape, shapeStyle);
                    }
                }

                // 元の図形を削除
                ComExceptionHandler.ExecuteComOperation(
                    () => originalShape.Delete(),
                    "元図形削除");
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("図形分割", ex);
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
                // 入力検証
                ErrorHandler.ValidateShapes(originalShapes, Constants.MIN_SHAPES_FOR_DIVISION, "グリッド分割");
                ErrorHandler.ValidateRange(rows, Constants.MIN_ROWS, Constants.MAX_ROWS, "行数", "グリッド分割");
                ErrorHandler.ValidateRange(columns, Constants.MIN_COLUMNS, Constants.MAX_COLUMNS, "列数", "グリッド分割");
                ErrorHandler.ValidateRange(horizontalMargin, Constants.MIN_MARGIN, Constants.MAX_MARGIN, "水平マージン", "グリッド分割");
                ErrorHandler.ValidateRange(verticalMargin, Constants.MIN_MARGIN, Constants.MAX_MARGIN, "垂直マージン", "グリッド分割");

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

                // 5. セルサイズの検証
                ErrorHandler.ValidateCellSize(cellWidth, cellHeight, "グリッド分割");

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

                ComExceptionHandler.LogDebug($"グリッド分割完了: {rows}×{columns}, " +
                    $"範囲: {bounds.Width:F2}×{bounds.Height:F2}, 作成図形数: {createdShapes.Count}");
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("グリッド分割", ex);
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
                var info = ComExceptionHandler.HandleComOperation(
                    () => {
                        var shapeInfo = new ShapeInfo
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
                            shapeInfo.FillColor = shape.Fill.Visible == Office.MsoTriState.msoTrue ?
                                (int?)shape.Fill.ForeColor.RGB : null;
                            shapeInfo.FillTransparency = shape.Fill.Transparency;
                            shapeInfo.LineColor = shape.Line.Visible == Office.MsoTriState.msoTrue ?
                                (int?)shape.Line.ForeColor.RGB : null;
                            shapeInfo.LineWeight = shape.Line.Weight;
                            shapeInfo.LineDashStyle = shape.Line.DashStyle;
                        }
                        catch
                        {
                            // スタイル情報の取得に失敗した場合はデフォルト値を使用
                            shapeInfo.FillColor = Constants.DEFAULT_FILL_COLOR;
                            shapeInfo.LineColor = Constants.DEFAULT_LINE_COLOR;
                            shapeInfo.LineWeight = Constants.DEFAULT_LINE_WEIGHT;
                            shapeInfo.LineDashStyle = Office.MsoLineDashStyle.msoLineSolid;
                        }

                        return shapeInfo;
                    },
                    $"図形情報取得: {shape.Name}",
                    throwOnError: false);

                if (info != null)
                {
                    shapeInfos.Add(info);
                    ComExceptionHandler.LogDebug($"図形情報取得成功: {info.Name} ({info.Left:F1}, {info.Top:F1}, {info.Width:F1}×{info.Height:F1})");
                }
            }

            ComExceptionHandler.LogDebug($"取得した図形情報数: {shapeInfos.Count}/{shapes.Count}");
            return shapeInfos;
        }

        /// <summary>
        /// 有効なスライドを取得
        /// </summary>
        private PowerPoint.Slide GetValidSlide(List<PowerPoint.Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                var slide = ComExceptionHandler.HandleComOperation(
                    () => {
                        var slideRef = shape.Parent as PowerPoint.Slide;
                        if (slideRef != null)
                        {
                            // スライドの有効性をテスト
                            var testName = slideRef.Name;
                            return slideRef;
                        }
                        return null;
                    },
                    "スライド取得",
                    throwOnError: false);

                if (slide != null)
                {
                    return slide;
                }
            }
            return null;
        }

        /// <summary>
        /// 境界を計算（事前取得した情報を使用）
        /// </summary>
        private ShapeGroupBounds CalculateBounds(List<ShapeInfo> shapeInfos)
        {
            return ShapeGroupBounds.FromShapeInfos(shapeInfos);
        }

        /// <summary>
        /// 事前取得した情報からスタイルを作成
        /// </summary>
        private ShapeStyle ExtractShapeStyleFromInfo(ShapeInfo shapeInfo)
        {
            return new ShapeStyle
            {
                FillColor = shapeInfo.FillColor ?? Constants.DEFAULT_FILL_COLOR,
                FillTransparency = shapeInfo.FillTransparency,
                LineColor = shapeInfo.LineColor ?? Constants.DEFAULT_LINE_COLOR,
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
                    float left = bounds.Left + col * (cellWidth + horizontalMargin);
                    float top = bounds.Top + row * (cellHeight + verticalMargin);

                    // 座標の有効性をチェック
                    if (!ErrorHandler.ValidateCoordinates(left, top, "グリッド図形作成"))
                    {
                        continue;
                    }

                    var newShape = ComExceptionHandler.HandleComOperation(
                        () => slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            left, top, cellWidth, cellHeight),
                        $"図形作成 (行{row + 1}, 列{col + 1})",
                        throwOnError: false);

                    if (newShape != null)
                    {
                        createdShapes.Add(newShape);
                        ApplyShapeStyleSafe(newShape, style);
                        ComExceptionHandler.LogDebug($"図形作成成功: 行{row + 1}列{col + 1} ({left:F1}, {top:F1})");
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
                var success = ComExceptionHandler.HandleComOperation(
                    () => shape.Delete(),
                    "図形削除",
                    throwOnError: false);

                if (success)
                {
                    deletedCount++;
                }
            }
            ComExceptionHandler.LogDebug($"削除した図形数: {deletedCount}/{shapes.Count}");
        }

        #endregion

        #region 共通スタイル処理メソッド

        /// <summary>
        /// 図形のスタイルを抽出する
        /// </summary>
        private ShapeStyle ExtractShapeStyle(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var style = new ShapeStyle();

                    // 塗りつぶし情報
                    if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        style.FillColor = shape.Fill.ForeColor.RGB;
                        style.FillTransparency = shape.Fill.Transparency;
                    }
                    else
                    {
                        style.FillColor = Constants.DEFAULT_FILL_COLOR;
                        style.FillTransparency = Constants.DEFAULT_TRANSPARENCY;
                    }

                    // 線の情報
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        style.LineColor = shape.Line.ForeColor.RGB;
                        style.LineWeight = shape.Line.Weight;
                        style.LineDashStyle = shape.Line.DashStyle;
                    }
                    else
                    {
                        style.LineColor = Constants.DEFAULT_LINE_COLOR;
                        style.LineWeight = Constants.DEFAULT_LINE_WEIGHT;
                        style.LineDashStyle = Office.MsoLineDashStyle.msoLineSolid;
                    }

                    // 影の情報
                    if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                    {
                        style.HasShadow = true;
                        style.ShadowColor = shape.Shadow.ForeColor.RGB;
                    }

                    return style;
                },
                "スタイル抽出",
                ShapeStyle.CreateDefault());
        }

        /// <summary>
        /// 図形にスタイルを適用する
        /// </summary>
        private void ApplyShapeStyle(PowerPoint.Shape shape, ShapeStyle style)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
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
                },
                "スタイル適用");
        }

        /// <summary>
        /// 安全な図形スタイル適用
        /// </summary>
        private void ApplyShapeStyleSafe(PowerPoint.Shape shape, ShapeStyle style)
        {
            // 塗りつぶしを適用
            if (style.FillColor.HasValue)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () => {
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.ForeColor.RGB = style.FillColor.Value;
                        shape.Fill.Transparency = style.FillTransparency;
                    },
                    "塗りつぶし適用",
                    suppressErrors: true);
            }

            // 線を適用
            if (style.LineColor.HasValue)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () => {
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.ForeColor.RGB = style.LineColor.Value;
                        shape.Line.Weight = style.LineWeight;
                        shape.Line.DashStyle = style.LineDashStyle;
                    },
                    "線適用",
                    suppressErrors: true);
            }

            // 影を適用
            if (style.HasShadow)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () => {
                        shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                        shape.Shadow.ForeColor.RGB = style.ShadowColor;
                    },
                    "影適用",
                    suppressErrors: true);
            }
        }

        #endregion
    }
}