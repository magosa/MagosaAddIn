using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形複製・配列機能を提供するクラス
    /// </summary>
    public class ShapeArrayer
    {
        /// <summary>
        /// 線形配列
        /// 指定方向・個数・間隔で図形を複製配列
        /// </summary>
        public void LinearArray(List<PowerPoint.Shape> shapes, LinearArrayOptions options)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ARRAY, "線形配列");
            ErrorHandler.ValidateRange(options.Count, Constants.MIN_ARRAY_COUNT, Constants.MAX_ARRAY_COUNT, 
                "個数", "線形配列");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                ComExceptionHandler.LogDebug($"線形配列: 角度={options.Angle:F1}°, 個数={options.Count}, 間隔={options.Spacing:F1}pt");

                // 角度をラジアンに変換
                double angleRad = options.Angle * Math.PI / 180.0;
                float dx = (float)(Math.Cos(angleRad) * options.Spacing);
                float dy = (float)(Math.Sin(angleRad) * options.Spacing);

                foreach (var shape in shapes)
                {
                    for (int i = 1; i < options.Count; i++)
                    {
                        var newShape = DuplicateShape(shape);
                        newShape.Left = shape.Left + dx * i;
                        newShape.Top = shape.Top + dy * i;
                    }
                }
            }, "線形配列");
        }

        /// <summary>
        /// 円形配列（回転コピー統合版）
        /// 等分配置モードまたは角度指定モードで円形/回転配列を実行
        /// </summary>
        public void CircularArray(List<PowerPoint.Shape> shapes, CircularArrayOptions options)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ARRAY, "円形配列");
            ErrorHandler.ValidateRange(options.Count, Constants.MIN_ARRAY_COUNT, Constants.MAX_ARRAY_COUNT, 
                "個数", "円形配列");

            // 等分配置モードの場合のみ半径を検証
            if (options.AngleMode == ArrayAngleMode.EqualDivision)
            {
                ErrorHandler.ValidateRange(options.Radius, Constants.MIN_RADIUS, Constants.MAX_RADIUS, 
                    "半径", "円形配列");
            }

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                string modeText = options.AngleMode == ArrayAngleMode.EqualDivision ? "等分配置" : "角度指定";
                ComExceptionHandler.LogDebug($"円形配列({modeText}): 個数={options.Count}, " +
                    $"中心指定={options.CenterSource}, 図形回転={options.RotateShapes}");

                // 回転中心を事前に決定（全図形で共通）
                float centerX, centerY;
                switch (options.CenterSource)
                {
                    case CenterSource.TargetShapeCenter:
                        // 最初に選択した図形の中心を使用
                        centerX = shapes[0].Left + shapes[0].Width / 2;
                        centerY = shapes[0].Top + shapes[0].Height / 2;
                        ComExceptionHandler.LogDebug($"回転中心: 最初の図形の中心 ({centerX:F1}, {centerY:F1})");
                        break;
                    default: // CustomCoordinate
                        centerX = options.CenterX;
                        centerY = options.CenterY;
                        ComExceptionHandler.LogDebug($"回転中心: カスタム座標 ({centerX:F1}, {centerY:F1})");
                        break;
                }

                foreach (var shape in shapes)
                {

                    // 角度ステップを計算
                    float angleStep = options.AngleMode == ArrayAngleMode.EqualDivision
                        ? 360.0f / options.Count      // 等分配置
                        : options.AngleStep;          // 角度指定

                    // 等分配置モードの場合の配置処理
                    if (options.AngleMode == ArrayAngleMode.EqualDivision)
                    {
                        // 元の図形 + 複製N個 = 合計N+1個
                        // 0からN-1までループして、N個の複製を作成
                        for (int i = 0; i < options.Count; i++)
                        {
                            float angle = options.StartAngle + angleStep * i;
                            double angleRad = angle * Math.PI / 180.0;

                            float newCenterX = centerX + (float)(Math.Cos(angleRad) * options.Radius);
                            float newCenterY = centerY + (float)(Math.Sin(angleRad) * options.Radius);

                            var newShape = DuplicateShape(shape);
                            newShape.Left = newCenterX - newShape.Width / 2;
                            newShape.Top = newCenterY - newShape.Height / 2;

                            if (options.RotateShapes)
                            {
                                newShape.Rotation = angle + 90; // 接線方向に回転
                            }
                        }
                    }
                    else // 角度指定モード（回転コピー）
                    {
                        // 元の図形の中心からの相対位置
                        float shapeCenterX = shape.Left + shape.Width / 2;
                        float shapeCenterY = shape.Top + shape.Height / 2;
                        float relX = shapeCenterX - centerX;
                        float relY = shapeCenterY - centerY;
                        float distance = (float)Math.Sqrt(relX * relX + relY * relY);
                        float initialAngle = (float)Math.Atan2(relY, relX);

                        for (int i = 1; i < options.Count; i++)
                        {
                            float rotationAngle = options.StartAngle + angleStep * i;
                            double rotationRad = rotationAngle * Math.PI / 180.0;
                            float newAngle = initialAngle + (float)rotationRad;

                            float newCenterX = centerX + (float)(Math.Cos(newAngle) * distance);
                            float newCenterY = centerY + (float)(Math.Sin(newAngle) * distance);

                            var newShape = DuplicateShape(shape);
                            newShape.Left = newCenterX - newShape.Width / 2;
                            newShape.Top = newCenterY - newShape.Height / 2;

                            if (options.RotateShapes)
                            {
                                newShape.Rotation = shape.Rotation + rotationAngle;
                            }
                        }
                    }
                }
            }, "円形配列");
        }

        /// <summary>
        /// グリッド配列（角度対応・線形配列統合版）
        /// 行×列で格子状に配列、角度を指定してグリッド全体を回転可能
        /// </summary>
        public void GridArray(List<PowerPoint.Shape> shapes, GridArrayOptions options)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ARRAY, "グリッド配列");
            ErrorHandler.ValidateRange(options.Rows, Constants.MIN_ROWS, Constants.MAX_ROWS, 
                "行数", "グリッド配列");
            ErrorHandler.ValidateRange(options.Columns, Constants.MIN_COLUMNS, Constants.MAX_COLUMNS, 
                "列数", "グリッド配列");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                ComExceptionHandler.LogDebug($"グリッド配列: {options.Rows}行×{options.Columns}列, " +
                    $"間隔=({options.HorizontalSpacing:F1}, {options.VerticalSpacing:F1})pt, 角度={options.Angle:F1}°");

                // 角度をラジアンに変換
                double angleRad = options.Angle * Math.PI / 180.0;
                double cosAngle = Math.Cos(angleRad);
                double sinAngle = Math.Sin(angleRad);

                foreach (var shape in shapes)
                {
                    // 元の図形の中心座標を回転中心として使用
                    float baseCenterX = shape.Left + shape.Width / 2;
                    float baseCenterY = shape.Top + shape.Height / 2;

                    for (int row = 0; row < options.Rows; row++)
                    {
                        for (int col = 0; col < options.Columns; col++)
                        {
                            // 最初の位置（0,0）は元の図形なのでスキップ
                            if (row == 0 && col == 0) continue;

                            // グリッド座標系での相対位置
                            float gridX = (shape.Width + options.HorizontalSpacing) * col;
                            float gridY = (shape.Height + options.VerticalSpacing) * row;

                            // 回転行列を適用
                            float rotatedX = (float)(gridX * cosAngle - gridY * sinAngle);
                            float rotatedY = (float)(gridX * sinAngle + gridY * cosAngle);

                            // 新しい図形を作成して配置
                            var newShape = DuplicateShape(shape);
                            float newCenterX = baseCenterX + rotatedX;
                            float newCenterY = baseCenterY + rotatedY;
                            newShape.Left = newCenterX - newShape.Width / 2;
                            newShape.Top = newCenterY - newShape.Height / 2;
                        }
                    }
                }
            }, "グリッド配列");
        }

        /// <summary>
        /// パス配列
        /// カスタムパス（線）に沿って配列
        /// </summary>
        public void PathArray(PowerPoint.Shape shape, PowerPoint.Shape pathShape, PathArrayOptions options)
        {
            if (shape == null)
                throw new ArgumentNullException(nameof(shape), "配列する図形が指定されていません。");
            if (pathShape == null)
                throw new ArgumentNullException(nameof(pathShape), "パス図形が指定されていません。");

            ErrorHandler.ValidateRange(options.Count, Constants.MIN_ARRAY_COUNT, Constants.MAX_ARRAY_COUNT, 
                "個数", "パス配列");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                ComExceptionHandler.LogDebug($"パス配列: 個数={options.Count}, " +
                    $"等間隔={options.EqualSpacing}, 回転={options.RotateAlongPath}");

                // パスに沿った点を取得
                var points = GetPathPoints(pathShape, options.Count);

                for (int i = 1; i < options.Count; i++)
                {
                    var newShape = DuplicateShape(shape);
                    var (x, y) = points[i];

                    newShape.Left = x - newShape.Width / 2;
                    newShape.Top = y - newShape.Height / 2;

                    // パスに沿って回転する場合
                    if (options.RotateAlongPath && i > 0)
                    {
                        float angle = GetPathAngleAt(points, i);
                        newShape.Rotation = angle;
                    }
                }
            }, "パス配列");
        }

        /// <summary>
        /// 回転コピー
        /// 指定角度ずつ回転させながら複製
        /// </summary>
        public void RotationCopy(List<PowerPoint.Shape> shapes, RotationCopyOptions options)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ARRAY, "回転コピー");
            ErrorHandler.ValidateRange(options.Count, Constants.MIN_ARRAY_COUNT, Constants.MAX_ARRAY_COUNT, 
                "個数", "回転コピー");
            ErrorHandler.ValidateRange(options.Angle, Constants.MIN_ROTATION_ANGLE, Constants.MAX_ROTATION_ANGLE, 
                "角度", "回転コピー");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                ComExceptionHandler.LogDebug($"回転コピー: 角度={options.Angle:F1}°, 個数={options.Count}");

                foreach (var shape in shapes)
                {
                    // 回転中心を決定
                    float centerX, centerY;
                    if (options.UseShapeCenter)
                    {
                        centerX = shape.Left + shape.Width / 2;
                        centerY = shape.Top + shape.Height / 2;
                    }
                    else
                    {
                        centerX = options.CenterX;
                        centerY = options.CenterY;
                    }

                    // 元の図形の中心からの相対位置
                    float shapeCenterX = shape.Left + shape.Width / 2;
                    float shapeCenterY = shape.Top + shape.Height / 2;
                    float relX = shapeCenterX - centerX;
                    float relY = shapeCenterY - centerY;
                    float distance = (float)Math.Sqrt(relX * relX + relY * relY);
                    float initialAngle = (float)Math.Atan2(relY, relX);

                    for (int i = 1; i < options.Count; i++)
                    {
                        float rotationAngle = options.Angle * i;
                        double rotationRad = rotationAngle * Math.PI / 180.0;
                        float newAngle = initialAngle + (float)rotationRad;

                        float newCenterX = centerX + (float)(Math.Cos(newAngle) * distance);
                        float newCenterY = centerY + (float)(Math.Sin(newAngle) * distance);

                        var newShape = DuplicateShape(shape);
                        newShape.Left = newCenterX - newShape.Width / 2;
                        newShape.Top = newCenterY - newShape.Height / 2;
                        newShape.Rotation = shape.Rotation + rotationAngle;
                    }
                }
            }, "回転コピー");
        }

        #region プライベートヘルパーメソッド

        /// <summary>
        /// 図形を複製
        /// </summary>
        private PowerPoint.Shape DuplicateShape(PowerPoint.Shape shape)
        {
            var duplicated = shape.Duplicate();
            return duplicated[1];
        }

        /// <summary>
        /// パス図形から点列を取得
        /// </summary>
        private List<(float X, float Y)> GetPathPoints(PowerPoint.Shape pathShape, int count)
        {
            var points = new List<(float X, float Y)>();

            // 簡易実装: 直線の場合の処理
            // パス図形の始点と終点を結ぶ直線上に等間隔で配置
            float startX = pathShape.Left;
            float startY = pathShape.Top;
            float endX = pathShape.Left + pathShape.Width;
            float endY = pathShape.Top + pathShape.Height;

            for (int i = 0; i < count; i++)
            {
                float t = (float)i / (count - 1);
                float x = startX + (endX - startX) * t;
                float y = startY + (endY - startY) * t;
                points.Add((x, y));
            }

            ComExceptionHandler.LogDebug($"パス点取得: {points.Count}個の点を生成");
            return points;
        }

        /// <summary>
        /// パスの指定位置での角度を取得
        /// </summary>
        private float GetPathAngleAt(List<(float X, float Y)> points, int index)
        {
            if (index <= 0 || index >= points.Count) return 0;

            var (x1, y1) = points[index - 1];
            var (x2, y2) = points[index];

            float dx = x2 - x1;
            float dy = y2 - y1;

            float angleRad = (float)Math.Atan2(dy, dx);
            float angleDeg = angleRad * 180.0f / (float)Math.PI;

            return angleDeg;
        }

        #endregion
    }
}
