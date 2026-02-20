using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形サイズ一括調整機能を提供するクラス
    /// </summary>
    public class ShapeResizer
    {
        /// <summary>
        /// 基準図形サイズ適用
        /// 最初の図形を基準に、他の図形を同じサイズに調整
        /// </summary>
        public void ResizeToReference(List<PowerPoint.Shape> shapes, ResizeMode mode = ResizeMode.KeepCenter)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "基準図形サイズ適用");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                var referenceShape = shapes[0];
                float refWidth = referenceShape.Width;
                float refHeight = referenceShape.Height;

                ComExceptionHandler.LogDebug($"基準図形サイズ適用: {refWidth:F1} x {refHeight:F1}pt");

                for (int i = 1; i < shapes.Count; i++)
                {
                    ResizeShape(shapes[i], refWidth, refHeight, mode);
                }
            }, "基準図形サイズ適用");
        }

        /// <summary>
        /// 比率保持リサイズ（幅統一）
        /// 基準図形の幅に合わせ、アスペクト比を保って高さを調整
        /// </summary>
        public void ResizeToWidthKeepRatio(List<PowerPoint.Shape> shapes)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "幅統一・比率保持");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                var referenceShape = shapes[0];
                float refWidth = referenceShape.Width;

                ComExceptionHandler.LogDebug($"幅統一・比率保持: {refWidth:F1}pt");

                for (int i = 1; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    float aspectRatio = shape.Height / shape.Width;
                    float newHeight = refWidth * aspectRatio;

                    ResizeShape(shape, refWidth, newHeight, ResizeMode.KeepCenter);
                }
            }, "幅統一・比率保持");
        }

        /// <summary>
        /// 比率保持リサイズ（高さ統一）
        /// 基準図形の高さに合わせ、アスペクト比を保って幅を調整
        /// </summary>
        public void ResizeToHeightKeepRatio(List<PowerPoint.Shape> shapes)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_ALIGNMENT, "高さ統一・比率保持");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                var referenceShape = shapes[0];
                float refHeight = referenceShape.Height;

                ComExceptionHandler.LogDebug($"高さ統一・比率保持: {refHeight:F1}pt");

                for (int i = 1; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    float aspectRatio = shape.Width / shape.Height;
                    float newWidth = refHeight * aspectRatio;

                    ResizeShape(shape, newWidth, refHeight, ResizeMode.KeepCenter);
                }
            }, "高さ統一・比率保持");
        }

        /// <summary>
        /// 最大サイズ統一
        /// 選択図形の中で最も大きい図形のサイズに統一
        /// </summary>
        public void ResizeToMaximum(List<PowerPoint.Shape> shapes, ResizeMode mode = ResizeMode.KeepCenter)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_RESIZE, "最大サイズ統一");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                // 面積で最大の図形を探す
                var maxShape = shapes.OrderByDescending(s => s.Width * s.Height).First();
                float maxWidth = maxShape.Width;
                float maxHeight = maxShape.Height;

                ComExceptionHandler.LogDebug($"最大サイズ統一: {maxWidth:F1} x {maxHeight:F1}pt");

                foreach (var shape in shapes)
                {
                    if (shape != maxShape)
                    {
                        ResizeShape(shape, maxWidth, maxHeight, mode);
                    }
                }
            }, "最大サイズ統一");
        }

        /// <summary>
        /// 最小サイズ統一
        /// 選択図形の中で最も小さい図形のサイズに統一
        /// </summary>
        public void ResizeToMinimum(List<PowerPoint.Shape> shapes, ResizeMode mode = ResizeMode.KeepCenter)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_RESIZE, "最小サイズ統一");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                // 面積で最小の図形を探す
                var minShape = shapes.OrderBy(s => s.Width * s.Height).First();
                float minWidth = minShape.Width;
                float minHeight = minShape.Height;

                ComExceptionHandler.LogDebug($"最小サイズ統一: {minWidth:F1} x {minHeight:F1}pt");

                foreach (var shape in shapes)
                {
                    if (shape != minShape)
                    {
                        ResizeShape(shape, minWidth, minHeight, mode);
                    }
                }
            }, "最小サイズ統一");
        }

        /// <summary>
        /// パーセント拡大縮小
        /// 全図形を指定%で拡大縮小（中心位置保持）
        /// </summary>
        public void ResizeByPercentage(List<PowerPoint.Shape> shapes, float percentage)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_RESIZE, "パーセント拡大縮小");
            ErrorHandler.ValidateRange(percentage, Constants.MIN_PERCENTAGE, Constants.MAX_PERCENTAGE, 
                "パーセント", "パーセント拡大縮小");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                float scale = percentage / 100.0f;
                ComExceptionHandler.LogDebug($"パーセント拡大縮小: {percentage:F1}%（倍率: {scale:F3}）");

                foreach (var shape in shapes)
                {
                    float newWidth = shape.Width * scale;
                    float newHeight = shape.Height * scale;

                    ResizeShape(shape, newWidth, newHeight, ResizeMode.KeepCenter);
                }
            }, "パーセント拡大縮小");
        }

        /// <summary>
        /// 固定サイズ設定
        /// 直接mm/cm/pt単位でサイズ指定
        /// </summary>
        public void ResizeToFixedSize(List<PowerPoint.Shape> shapes, float width, float height, 
            SizeUnit unit, bool keepRatio = false, ResizeMode mode = ResizeMode.KeepCenter)
        {
            ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_RESIZE, "固定サイズ設定");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                // 単位をptに変換
                float widthPt = ConvertToPoints(width, unit);
                float heightPt = ConvertToPoints(height, unit);

                ErrorHandler.ValidateRange(widthPt, Constants.MIN_SHAPE_WIDTH, Constants.MAX_COORDINATE, 
                    "幅", "固定サイズ設定");
                ErrorHandler.ValidateRange(heightPt, Constants.MIN_SHAPE_HEIGHT, Constants.MAX_COORDINATE, 
                    "高さ", "固定サイズ設定");

                ComExceptionHandler.LogDebug($"固定サイズ設定: {widthPt:F1} x {heightPt:F1}pt（比率保持: {keepRatio}）");

                foreach (var shape in shapes)
                {
                    if (keepRatio)
                    {
                        // アスペクト比保持：幅基準でリサイズ
                        float aspectRatio = shape.Height / shape.Width;
                        float calculatedHeight = widthPt * aspectRatio;
                        ResizeShape(shape, widthPt, calculatedHeight, mode);
                    }
                    else
                    {
                        ResizeShape(shape, widthPt, heightPt, mode);
                    }
                }
            }, "固定サイズ設定");
        }

        #region プライベートヘルパーメソッド

        /// <summary>
        /// 図形をリサイズする（中心位置または左上位置を保持）
        /// </summary>
        private void ResizeShape(PowerPoint.Shape shape, float newWidth, float newHeight, ResizeMode mode)
        {
            if (mode == ResizeMode.KeepCenter)
            {
                // 中心位置を保持
                float centerX = shape.Left + shape.Width / 2;
                float centerY = shape.Top + shape.Height / 2;

                shape.Width = newWidth;
                shape.Height = newHeight;

                shape.Left = centerX - newWidth / 2;
                shape.Top = centerY - newHeight / 2;
            }
            else
            {
                // 左上位置を保持
                shape.Width = newWidth;
                shape.Height = newHeight;
            }
        }

        /// <summary>
        /// 単位をポイントに変換
        /// </summary>
        private float ConvertToPoints(float value, SizeUnit unit)
        {
            switch (unit)
            {
                case SizeUnit.Point:
                    return value;
                case SizeUnit.Millimeter:
                    return value * Constants.MM_TO_PT;
                case SizeUnit.Centimeter:
                    return value * Constants.CM_TO_PT;
                default:
                    throw new ArgumentException($"未対応の単位: {unit}");
            }
        }

        /// <summary>
        /// ポイントを指定単位に変換
        /// </summary>
        public float ConvertFromPoints(float valuePt, SizeUnit unit)
        {
            switch (unit)
            {
                case SizeUnit.Point:
                    return valuePt;
                case SizeUnit.Millimeter:
                    return valuePt * Constants.PT_TO_MM;
                case SizeUnit.Centimeter:
                    return valuePt * Constants.PT_TO_CM;
                default:
                    throw new ArgumentException($"未対応の単位: {unit}");
            }
        }

        #endregion
    }
}
