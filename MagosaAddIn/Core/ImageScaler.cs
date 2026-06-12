using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 画像倍率同期スケーリング機能を提供するクラス。
    /// 実寸法を考慮した倍率計算に対応。
    /// </summary>
    public class ImageScaler
    {
        #region 公開メソッド

        /// <summary>
        /// 2点間の距離を計算（ピクセル）
        /// </summary>
        /// <param name="x1">起点X座標</param>
        /// <param name="y1">起点Y座標</param>
        /// <param name="x2">終点X座標</param>
        /// <param name="y2">終点Y座標</param>
        /// <param name="mode">測定方向</param>
        /// <returns>距離（ピクセル）</returns>
        public float CalculateDistance(float x1, float y1, float x2, float y2,
            MeasurementMode mode = MeasurementMode.Free)
        {
            switch (mode)
            {
                case MeasurementMode.HorizontalOnly:
                    return Math.Abs(x2 - x1);
                case MeasurementMode.VerticalOnly:
                    return Math.Abs(y2 - y1);
                default:
                    float dx = x2 - x1;
                    float dy = y2 - y1;
                    return (float)Math.Sqrt(dx * dx + dy * dy);
            }
        }

        /// <summary>
        /// 画像の倍率を計算（実寸法 / ピクセル長）
        /// 単位は内部でmmに正規化してから計算
        /// </summary>
        /// <param name="realLength">実寸法の値</param>
        /// <param name="unit">実寸法の単位</param>
        /// <param name="pixelLength">測定区間のピクセル長</param>
        /// <returns>倍率（mm/px）</returns>
        public float CalculateImageRatio(float realLength, SizeUnit unit, float pixelLength)
        {
            if (pixelLength <= 0f)
                throw new ArgumentException("ピクセル長は0より大きい値である必要があります。", nameof(pixelLength));

            float realLengthMm = NormalizeRealLengthToMillimeter(realLength, unit);
            return realLengthMm / pixelLength;
        }

        /// <summary>
        /// 実寸法を内部基準単位（mm）へ正規化
        /// </summary>
        /// <param name="value">変換前の値</param>
        /// <param name="unit">入力単位</param>
        /// <returns>ミリメートル換算値</returns>
        public float NormalizeRealLengthToMillimeter(float value, SizeUnit unit)
        {
            switch (unit)
            {
                case SizeUnit.Millimeter:
                    return value;
                case SizeUnit.Centimeter:
                    return value * 10f;
                case SizeUnit.Point:
                    // 1pt = 25.4/72 mm
                    return value * (25.4f / 72f);
                default:
                    throw new ArgumentException($"未対応の単位: {unit}", nameof(unit));
            }
        }

        /// <summary>
        /// 2つの画像の実寸法を考慮したスケール係数を計算。
        /// ScaleFactor = Ratio1 / Ratio2 = (L1real × L2px) / (L2real × L1px)
        /// </summary>
        /// <param name="image1">基準画像</param>
        /// <param name="x1Start">画像①起点X（ピクセル）</param>
        /// <param name="y1Start">画像①起点Y（ピクセル）</param>
        /// <param name="x1End">画像①終点X（ピクセル）</param>
        /// <param name="y1End">画像①終点Y（ピクセル）</param>
        /// <param name="image1RealLength">画像①の測定部分の実寸法</param>
        /// <param name="image1Unit">画像①の実寸法単位</param>
        /// <param name="image2">対象画像</param>
        /// <param name="x2Start">画像②起点X（ピクセル）</param>
        /// <param name="y2Start">画像②起点Y（ピクセル）</param>
        /// <param name="x2End">画像②終点X（ピクセル）</param>
        /// <param name="y2End">画像②終点Y（ピクセル）</param>
        /// <param name="image2RealLength">画像②の測定部分の実寸法</param>
        /// <param name="image2Unit">画像②の実寸法単位</param>
        /// <param name="mode1">画像①の測定方向</param>
        /// <param name="mode2">画像②の測定方向</param>
        /// <param name="image1BitmapWidth">
        /// プレビュー用クリップボードビットマップの総ピクセル幅（画像①）。
        /// 0の場合は補正なし（スクリーン解像度と同等）。
        /// </param>
        /// <param name="image2BitmapWidth">
        /// プレビュー用クリップボードビットマップの総ピクセル幅（画像②）。
        /// 0の場合は補正なし（スクリーン解像度と同等）。
        /// </param>
        /// <returns>スケール係数（画像②に乗じる倍率）</returns>
        /// <remarks>
        /// 正しい計算式：F = (ratio2/ratio1) × (Px2_total × W1) / (Px1_total × W2)
        ///
        /// クリップボードがスクリーン解像度を返す場合 → Px ∝ W → 補正係数=1.0 → ratio2/ratio1 と等価
        /// クリップボードが原寸解像度を返す場合       → Px 独立 → 補正係数≠1.0 → 補正必要
        ///
        /// どちらの場合でも同じ式が正しく機能する。
        /// </remarks>
        public float CalculateScaleFactor(
            PowerPoint.Shape image1,
            float x1Start, float y1Start, float x1End, float y1End,
            float image1RealLength, SizeUnit image1Unit,
            PowerPoint.Shape image2,
            float x2Start, float y2Start, float x2End, float y2End,
            float image2RealLength, SizeUnit image2Unit,
            MeasurementMode mode1 = MeasurementMode.Free,
            MeasurementMode mode2 = MeasurementMode.Free,
            int image1BitmapWidth = 0,
            int image2BitmapWidth = 0)
        {
            float l1Px = CalculateDistance(x1Start, y1Start, x1End, y1End, mode1);
            float l2Px = CalculateDistance(x2Start, y2Start, x2End, y2End, mode2);

            if (l1Px < 0.001f)
                throw new ArgumentException("画像①の測定区間の長さが0です。異なる座標を指定してください。");
            if (l2Px < 0.001f)
                throw new ArgumentException("画像②の測定区間の長さが0です。異なる座標を指定してください。");

            float ratio1 = CalculateImageRatio(image1RealLength, image1Unit, l1Px);
            float ratio2 = CalculateImageRatio(image2RealLength, image2Unit, l2Px);

            if (Math.Abs(ratio1) < 1e-9f)
                throw new InvalidOperationException("画像①の倍率がゼロです。実寸法または座標を確認してください。");

            // 基本スケール係数: ratio2 / ratio1
            float scaleFactor = ratio2 / ratio1;

            // ビットマップ解像度補正:
            // F = (ratio2/ratio1) × (Px2_total × W1) / (Px1_total × W2)
            // クリップボードが原寸解像度を返す場合でも正しく動作させるための補正。
            // スクリーン解像度の場合は補正係数=1.0となり影響なし。
            if (image1BitmapWidth > 0 && image2BitmapWidth > 0)
            {
                float w1 = image1.Width;
                float w2 = image2.Width;
                if (w1 > 0.001f && w2 > 0.001f)
                {
                    float correctionFactor = ((float)image2BitmapWidth * w1)
                                           / ((float)image1BitmapWidth * w2);
                    scaleFactor *= correctionFactor;
                }
            }

            return scaleFactor;
        }

        /// <summary>
        /// 画像②にスケーリングを適用する
        /// </summary>
        /// <param name="image2">対象画像</param>
        /// <param name="scaleFactor">スケール係数</param>
        /// <param name="positionMode">リサイズ後の位置保持モード</param>
        public void ApplyScaling(PowerPoint.Shape image2, float scaleFactor,
            ResizeMode positionMode = ResizeMode.KeepCenter)
        {
            if (image2 == null)
                throw new ArgumentNullException(nameof(image2));
            if (scaleFactor <= 0f || float.IsNaN(scaleFactor) || float.IsInfinity(scaleFactor))
                throw new ArgumentException($"スケール係数が無効です: {scaleFactor}", nameof(scaleFactor));

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                float newWidth = image2.Width * scaleFactor;
                float newHeight = image2.Height * scaleFactor;

                const float MinSize = 1f;
                const float MaxSize = 10000f;
                if (newWidth < MinSize || newHeight < MinSize || newWidth > MaxSize || newHeight > MaxSize)
                    throw new InvalidOperationException(
                        $"計算結果が許可範囲外です（最小 {MinSize}pt、最大 {MaxSize}pt）。\n" +
                        $"計算後サイズ: {newWidth:F1} × {newHeight:F1} pt");

                if (positionMode == ResizeMode.KeepCenter)
                {
                    float centerX = image2.Left + image2.Width / 2f;
                    float centerY = image2.Top + image2.Height / 2f;

                    image2.Width = newWidth;
                    image2.Height = newHeight;

                    image2.Left = centerX - newWidth / 2f;
                    image2.Top = centerY - newHeight / 2f;
                }
                else
                {
                    // KeepTopLeft: Left/Top は変更しない
                    image2.Width = newWidth;
                    image2.Height = newHeight;
                }

                ComExceptionHandler.LogDebug(
                    $"画像スケーリング適用: {scaleFactor:F4}倍 → {newWidth:F1}×{newHeight:F1}pt ({positionMode})");
            }, "画像スケーリング適用");
        }

        /// <summary>
        /// 入力値の妥当性を検証する
        /// </summary>
        /// <returns>(IsValid, ErrorMessage) のタプル</returns>
        public (bool IsValid, string ErrorMessage) ValidateInputs(
            PowerPoint.Shape image1,
            float x1Start, float y1Start, float x1End, float y1End,
            float image1RealLength,
            PowerPoint.Shape image2,
            float x2Start, float y2Start, float x2End, float y2End,
            float image2RealLength,
            MeasurementMode mode = MeasurementMode.Free)
        {
            // 1. Null チェック
            if (image1 == null)
                return (false, "画像①が選択されていません。");
            if (image2 == null)
                return (false, "画像②が選択されていません。");

            // 2. 同一画像チェック
            if (image1.Name == image2.Name)
                return (false, "異なる画像を選択してください。");

            // 3. 画像タイプチェック
            if (image1.Type != Office.MsoShapeType.msoPicture)
                return (false, "画像①は画像オブジェクトである必要があります。");
            if (image2.Type != Office.MsoShapeType.msoPicture)
                return (false, "画像②は画像オブジェクトである必要があります。");

            // 4. 実寸法チェック
            if (image1RealLength <= 0f || float.IsNaN(image1RealLength))
                return (false, "画像①の実寸法は0より大きい値を指定してください。");
            if (image2RealLength <= 0f || float.IsNaN(image2RealLength))
                return (false, "画像②の実寸法は0より大きい値を指定してください。");

            // 5. 座標範囲チェック（画像①）
            var (valid1, msg1) = ValidateCoordinatesForImage(
                x1Start, y1Start, x1End, y1End,
                image1.Width, image1.Height, "画像①");
            if (!valid1) return (false, msg1);

            // 6. 座標範囲チェック（画像②）
            var (valid2, msg2) = ValidateCoordinatesForImage(
                x2Start, y2Start, x2End, y2End,
                image2.Width, image2.Height, "画像②");
            if (!valid2) return (false, msg2);

            // 7. ゼロ距離チェック
            float dist1 = CalculateDistance(x1Start, y1Start, x1End, y1End, mode);
            if (dist1 < 0.001f)
                return (false, "画像①の測定区間の長さが0です。異なる座標を指定してください。");

            float dist2 = CalculateDistance(x2Start, y2Start, x2End, y2End, mode);
            if (dist2 < 0.001f)
                return (false, "画像②の測定区間の長さが0です。異なる座標を指定してください。");

            return (true, string.Empty);
        }

        #endregion

        #region プライベートヘルパー

        private (bool IsValid, string ErrorMessage) ValidateCoordinatesForImage(
            float xStart, float yStart, float xEnd, float yEnd,
            float imageWidth, float imageHeight, string imageName)
        {
            if (float.IsNaN(xStart) || float.IsNaN(yStart) || float.IsNaN(xEnd) || float.IsNaN(yEnd) ||
                float.IsInfinity(xStart) || float.IsInfinity(yStart) ||
                float.IsInfinity(xEnd) || float.IsInfinity(yEnd))
                return (false, $"{imageName}の座標に無効な値（NaN/Infinity）が含まれています。");

            if (xStart < 0 || yStart < 0 || xEnd < 0 || yEnd < 0)
                return (false, $"{imageName}の座標が無効です。0以上の値を指定してください。");

            if (xStart > imageWidth || xEnd > imageWidth)
                return (false,
                    $"{imageName}のX座標が無効です。0以上 {imageWidth:F0}px 以下を指定してください。");

            if (yStart > imageHeight || yEnd > imageHeight)
                return (false,
                    $"{imageName}のY座標が無効です。0以上 {imageHeight:F0}px 以下を指定してください。");

            return (true, string.Empty);
        }

        #endregion
    }
}
