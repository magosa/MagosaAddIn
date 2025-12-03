using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MagosaAddIn.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// /// 図形の整列機能を提供するクラス
    /// /// </summary>
    public class ShapeAligner
    {
        #region 基準整列機能

        /// <summary>
        /// 1つめの選択オブジェクトを基準にその他のオブジェクトの左端を揃える
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignToLeft(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "左端揃え");

                    var baseShape = shapes[0];
                    float targetPosition = baseShape.Left;

                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        currentShape.Left = targetPosition;
                    }
                },
                "左端揃え"); // suppressErrors: false がデフォルト
        }

        /// <summary>
        /// 1つめの選択オブジェクトを基準にその他のオブジェクトの右端を揃える
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignToRight(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "右端揃え");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Left + baseShape.Width; // 基準図形の右端

                    // 2番目以降の図形の右端を基準図形の右端に合わせる
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        currentShape.Left = targetPosition - currentShape.Width;
                    }
                },
                "右端揃え");
        }

        /// <summary>
        /// 1つめの選択オブジェクトを基準にその他のオブジェクトの上端を揃える
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignToTop(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "上端揃え");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Top; // 基準図形の上端

                    // 2番目以降の図形の上端を基準図形の上端に合わせる
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        currentShape.Top = targetPosition;
                    }
                },
                "上端揃え");
        }

        /// <summary>
        /// 1つめの選択オブジェクトを基準にその他のオブジェクトの下端を揃える
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignToBottom(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "下端揃え");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Top + baseShape.Height; // 基準図形の下端

                    // 2番目以降の図形の下端を基準図形の下端に合わせる
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        currentShape.Top = targetPosition - currentShape.Height;
                    }
                },
                "下端揃え");
        }

        #endregion

        #region 隣接整列機能

        /// <summary>
        /// 一番目に選択したオブジェクトの左端にその他の選択オブジェクトの右端を合わせる
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignLeftToRight(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "左端→右端整列");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Left; // 基準図形の左端

                    // 2番目以降の図形を移動
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        // 現在の図形の右端を基準図形の左端に合わせる
                        currentShape.Left = targetPosition - currentShape.Width;
                    }
                },
                "左端→右端整列");
        }

        /// <summary>
        /// 一番目に選択したオブジェクトの右端にその他の選択オブジェクトの左端を合わせる
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignRightToLeft(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "右端→左端整列");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Left + baseShape.Width; // 基準図形の右端

                    // 2番目以降の図形を移動
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        // 現在の図形の左端を基準図形の右端に合わせる
                        currentShape.Left = targetPosition;
                    }
                },
                "右端→左端整列");
        }

        /// <summary>
        /// 一番目に選択したオブジェクトの上端にその他の選択オブジェクトの下端を合わせる
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignTopToBottom(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "上端→下端整列");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Top; // 基準図形の上端

                    // 2番目以降の図形を移動
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        // 現在の図形の下端を基準図形の上端に合わせる
                        currentShape.Top = targetPosition - currentShape.Height;
                    }
                },
                "上端→下端整列");
        }

        /// <summary>
        /// 一番目に選択したオブジェクトの下端にその他の選択オブジェクトの上端を合わせる
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        public void AlignBottomToTop(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "下端→上端整列");

                    var baseShape = shapes[0]; // 基準となる図形（一番目に選択）
                    float targetPosition = baseShape.Top + baseShape.Height; // 基準図形の下端

                    // 2番目以降の図形を移動
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        // 現在の図形の上端を基準図形の下端に合わせる
                        currentShape.Top = targetPosition;
                    }
                },
                "下端→上端整列");
        }

        #endregion

        #region 拡張整列機能

        /// <summary>
        /// 図形を水平方向に中央揃えして等間隔で配置
        /// </summary>
        /// <param name="shapes">選択された図形のリスト</param>
        public void AlignAndDistributeHorizontal(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "水平中央揃え・等間隔配置");

                    var baseShape = shapes[0];
                    float baseCenterY = baseShape.Top + (baseShape.Height / 2);

                    // まず全ての図形を基準図形の水平中央に揃える
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        currentShape.Top = baseCenterY - (currentShape.Height / 2);
                    }

                    // 2つの図形の場合は中央揃えのみで終了
                    if (shapes.Count == 2)
                    {
                        return;
                    }

                    // 3つ以上の場合は等間隔配置も実行
                    // 左端と右端の図形を特定
                    var leftmostShape = shapes.OrderBy(s => s.Left).First();
                    var rightmostShape = shapes.OrderBy(s => s.Left + s.Width).Last();

                    float totalWidth = (rightmostShape.Left + rightmostShape.Width) - leftmostShape.Left;
                    float totalShapeWidth = shapes.Sum(s => s.Width);
                    float totalGap = totalWidth - totalShapeWidth;
                    float gapBetweenShapes = totalGap / (shapes.Count - 1);

                    // 左から順に配置
                    var sortedShapes = shapes.OrderBy(s => s.Left).ToList();
                    float currentLeft = leftmostShape.Left;

                    for (int i = 0; i < sortedShapes.Count; i++)
                    {
                        if (i > 0) // 最初の図形は移動しない
                        {
                            sortedShapes[i].Left = currentLeft;
                        }
                        currentLeft += sortedShapes[i].Width + gapBetweenShapes;
                    }
                },
                "水平中央揃え・等間隔配置");
        }

        /// <summary>
        /// 図形を垂直方向に中央揃えして等間隔で配置
        /// </summary>
        /// <param name="shapes">選択された図形のリスト</param>
        public void AlignAndDistributeVertical(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "垂直中央揃え・等間隔配置");

                    var baseShape = shapes[0];
                    float baseCenterX = baseShape.Left + (baseShape.Width / 2);

                    // まず全ての図形を基準図形の垂直中央に揃える
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        currentShape.Left = baseCenterX - (currentShape.Width / 2);
                    }

                    // 2つの図形の場合は中央揃えのみで終了
                    if (shapes.Count == 2)
                    {
                        return;
                    }

                    // 3つ以上の場合は等間隔配置も実行
                    // 上端と下端の図形を特定
                    var topmostShape = shapes.OrderBy(s => s.Top).First();
                    var bottommostShape = shapes.OrderBy(s => s.Top + s.Height).Last();

                    float totalHeight = (bottommostShape.Top + bottommostShape.Height) - topmostShape.Top;
                    float totalShapeHeight = shapes.Sum(s => s.Height);
                    float totalGap = totalHeight - totalShapeHeight;
                    float gapBetweenShapes = totalGap / (shapes.Count - 1);

                    // 上から順に配置
                    var sortedShapes = shapes.OrderBy(s => s.Top).ToList();
                    float currentTop = topmostShape.Top;

                    for (int i = 0; i < sortedShapes.Count; i++)
                    {
                        if (i > 0) // 最初の図形は移動しない
                        {
                            sortedShapes[i].Top = currentTop;
                        }
                        currentTop += sortedShapes[i].Height + gapBetweenShapes;
                    }
                },
                "垂直中央揃え・等間隔配置");
        }

        /// <summary>
        /// はじめに選択した図形を基準に任意のマージンで水平方向に配置
        /// 基準図形の左右に元の位置関係を保って配置する
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        /// <param name="margin">図形間のマージン（pt）</param>
        public void ArrangeHorizontalWithMargin(List<PowerPoint.Shape> shapes, float margin)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "水平マージン配置");
                    ErrorHandler.ValidateRange(margin, Constants.MIN_MARGIN, Constants.MAX_MARGIN, "マージン", "水平マージン配置");

                    var baseShape = shapes[0]; // 基準図形（一番目に選択）
                    float baseCenterX = baseShape.Left + (baseShape.Width / 2);

                    // 基準図形以外の図形を左側と右側に分類
                    var leftShapes = new List<PowerPoint.Shape>();
                    var rightShapes = new List<PowerPoint.Shape>();

                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        float currentCenterX = currentShape.Left + (currentShape.Width / 2);

                        if (currentCenterX < baseCenterX)
                        {
                            leftShapes.Add(currentShape);
                        }
                        else
                        {
                            rightShapes.Add(currentShape);
                        }
                    }

                    // 左側の図形を基準図形からの距離順（近い順）でソート
                    leftShapes = leftShapes.OrderByDescending(s => s.Left + s.Width).ToList();

                    // 右側の図形を基準図形からの距離順（近い順）でソート
                    rightShapes = rightShapes.OrderBy(s => s.Left).ToList();

                    // 左側の図形を配置（基準図形の左端から左方向へ）
                    float currentLeftPosition = baseShape.Left - margin;
                    foreach (var shape in leftShapes)
                    {
                        shape.Left = currentLeftPosition - shape.Width;
                        currentLeftPosition = shape.Left - margin;
                    }

                    // 右側の図形を配置（基準図形の右端から右方向へ）
                    float currentRightPosition = baseShape.Left + baseShape.Width + margin;
                    foreach (var shape in rightShapes)
                    {
                        shape.Left = currentRightPosition;
                        currentRightPosition = shape.Left + shape.Width + margin;
                    }

                    ComExceptionHandler.LogDebug($"水平マージン配置完了: マージン {margin:F2}pt, " +
                        $"左側 {leftShapes.Count}個, 右側 {rightShapes.Count}個, 基準図形1個");
                },
                "水平マージン配置");
        }

        /// <summary>
        /// はじめに選択した図形を基準に任意のマージンで垂直方向に配置
        /// 基準図形の上下に元の位置関係を保って配置する
        /// </summary>
        /// <param name="shapes">選択された図形のリスト（最初の要素が基準図形）</param>
        /// <param name="margin">図形間のマージン（pt）</param>
        public void ArrangeVerticalWithMargin(List<PowerPoint.Shape> shapes, float margin)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "垂直マージン配置");
                    ErrorHandler.ValidateRange(margin, Constants.MIN_MARGIN, Constants.MAX_MARGIN, "マージン", "垂直マージン配置");

                    var baseShape = shapes[0]; // 基準図形（一番目に選択）
                    float baseCenterY = baseShape.Top + (baseShape.Height / 2);

                    // 基準図形以外の図形を上側と下側に分類
                    var topShapes = new List<PowerPoint.Shape>();
                    var bottomShapes = new List<PowerPoint.Shape>();

                    for (int i = 1; i < shapes.Count; i++)
                    {
                        var currentShape = shapes[i];
                        float currentCenterY = currentShape.Top + (currentShape.Height / 2);

                        if (currentCenterY < baseCenterY)
                        {
                            topShapes.Add(currentShape);
                        }
                        else
                        {
                            bottomShapes.Add(currentShape);
                        }
                    }

                    // 上側の図形を基準図形からの距離順（近い順）でソート
                    topShapes = topShapes.OrderByDescending(s => s.Top + s.Height).ToList();

                    // 下側の図形を基準図形からの距離順（近い順）でソート
                    bottomShapes = bottomShapes.OrderBy(s => s.Top).ToList();

                    // 上側の図形を配置（基準図形の上端から上方向へ）
                    float currentTopPosition = baseShape.Top - margin;
                    foreach (var shape in topShapes)
                    {
                        shape.Top = currentTopPosition - shape.Height;
                        currentTopPosition = shape.Top - margin;
                    }

                    // 下側の図形を配置（基準図形の下端から下方向へ）
                    float currentBottomPosition = baseShape.Top + baseShape.Height + margin;
                    foreach (var shape in bottomShapes)
                    {
                        shape.Top = currentBottomPosition;
                        currentBottomPosition = shape.Top + shape.Height + margin;
                    }

                    ComExceptionHandler.LogDebug($"垂直マージン配置完了: マージン {margin:F2}pt, " +
                        $"上側 {topShapes.Count}個, 下側 {bottomShapes.Count}個, 基準図形1個");
                },
                "垂直マージン配置");
        }

        /// <summary>
        /// 図形をグリッド状に配置
        /// </summary>
        /// <param name="shapes">選択された図形のリスト</param>
        /// <param name="columns">列数</param>
        /// <param name="horizontalSpacing">水平間隔</param>
        /// <param name="verticalSpacing">垂直間隔</param>
        public void ArrangeInGrid(List<PowerPoint.Shape> shapes, int columns, float horizontalSpacing, float verticalSpacing)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "グリッド配置");
                    ErrorHandler.ValidateRange(columns, Constants.MIN_COLUMNS, shapes.Count, "列数", "グリッド配置");
                    ErrorHandler.ValidateRange(horizontalSpacing, Constants.MIN_MARGIN, Constants.MAX_SPACING, "水平間隔", "グリッド配置");
                    ErrorHandler.ValidateRange(verticalSpacing, Constants.MIN_MARGIN, Constants.MAX_SPACING, "垂直間隔", "グリッド配置");

                    var baseShape = shapes[0];
                    float startLeft = baseShape.Left;
                    float startTop = baseShape.Top;

                    for (int i = 0; i < shapes.Count; i++)
                    {
                        int row = i / columns;
                        int col = i % columns;

                        float newLeft = startLeft + col * (shapes[0].Width + horizontalSpacing);
                        float newTop = startTop + row * (shapes[0].Height + verticalSpacing);

                        shapes[i].Left = newLeft;
                        shapes[i].Top = newTop;
                    }
                },
                "グリッド配置");
        }

        /// <summary>
        /// 図形を円形に配置
        /// </summary>
        /// <param name="shapes">選択された図形のリスト</param>
        /// <param name="centerX">円の中心X座標</param>
        /// <param name="centerY">円の中心Y座標</param>
        /// <param name="radius">半径</param>
        public void ArrangeInCircle(List<PowerPoint.Shape> shapes, float centerX, float centerY, float radius)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "円形配置");
                    ErrorHandler.ValidateRange(centerX, Constants.MIN_CENTER_COORDINATE, Constants.MAX_CENTER_COORDINATE, "中心X座標", "円形配置");
                    ErrorHandler.ValidateRange(centerY, Constants.MIN_CENTER_COORDINATE, Constants.MAX_CENTER_COORDINATE, "中心Y座標", "円形配置");
                    ErrorHandler.ValidateRange(radius, Constants.MIN_RADIUS, Constants.MAX_RADIUS, "半径", "円形配置");

                    double angleStep = 2 * Math.PI / shapes.Count;

                    for (int i = 0; i < shapes.Count; i++)
                    {
                        double angle = i * angleStep;
                        float x = centerX + (float)(radius * Math.Cos(angle)) - (shapes[i].Width / 2);
                        float y = centerY + (float)(radius * Math.Sin(angle)) - (shapes[i].Height / 2);

                        shapes[i].Left = x;
                        shapes[i].Top = y;
                    }
                },
                "円形配置");
        }

        #endregion

    }
}