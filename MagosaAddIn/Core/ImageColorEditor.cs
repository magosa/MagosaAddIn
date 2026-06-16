using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using ImageMagick;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// Magick.NET を使って画像図形に色編集を適用し、元図形と差し替えるクラス
    /// </summary>
    public class ImageColorEditor
    {
        // ──────────────────────────────────────────────────────────────────────────
        //  デバッグログ（%TEMP%\MagosaAddIn_debug.log）
        // ──────────────────────────────────────────────────────────────────────────
        private static readonly string _logPath =
            Path.Combine(Path.GetTempPath(), "MagosaAddIn_debug.log");

        private static void Log(string msg)
        {
            try
            {
                string line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}{Environment.NewLine}";
                File.AppendAllText(_logPath, line, Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine("MagosaDebug: " + msg);
            }
            catch { /* ログ失敗は無視 */ }
        }

        // ──────────────────────────────────────────────────────────────────────────
        //  Apply
        // ──────────────────────────────────────────────────────────────────────────
        /// <summary>
        /// 指定した画像図形に色編集設定を適用し、元図形と差し替える。
        /// </summary>
        public void Apply(PowerPoint.Shape shape, ImageColorSettings settings)
        {
            Log("=== Apply 開始 ===");

            // ── STEP 1: shape の基本情報を読む ───────────────────────────────────
            string shapeName;
            Office.MsoShapeType shapeType;
            float left, top, width, height;
            int zOrder;

            try
            {
                shapeName = shape.Name;
                shapeType = shape.Type;
                left   = shape.Left;
                top    = shape.Top;
                width  = shape.Width;
                height = shape.Height;
                zOrder = shape.ZOrderPosition;
                Log($"STEP1 OK: name={shapeName}, type={shapeType}, " +
                    $"L={left:F1} T={top:F1} W={width:F1} H={height:F1} Z={zOrder}");
            }
            catch (COMException comEx)
            {
                Log($"STEP1 FAIL (COMException): HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                throw new InvalidOperationException(
                    $"[STEP1 shape プロパティ読み取り] HRESULT=0x{comEx.HResult:X8}: {comEx.Message}", comEx);
            }

            // ── STEP 2: 親スライドを取得 ──────────────────────────────────────────
            // Undo 後に shape.Parent が 0x800A01A8 (Object deleted) を返すことがある。
            // その場合は ActiveWindow.View.Slide からフォールバック取得する。
            PowerPoint.Slide slide;
            try
            {
                slide = (PowerPoint.Slide)shape.Parent;
                Log($"STEP2 OK (shape.Parent): slide={slide.SlideIndex}");
            }
            catch (COMException comEx)
            {
                Log($"STEP2 shape.Parent 失敗 (0x{comEx.HResult:X8}), ActiveWindow にフォールバック");
                try
                {
                    slide = (PowerPoint.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                    Log($"STEP2 Fallback OK: slide={slide.SlideIndex}");
                }
                catch (COMException comEx2)
                {
                    Log($"STEP2 Fallback FAIL: 0x{comEx2.HResult:X8} {comEx2.Message}");
                    throw new InvalidOperationException(
                        $"[STEP2 スライド取得失敗] shape.Parent:0x{comEx.HResult:X8}, ActiveWindow:0x{comEx2.HResult:X8}: {comEx2.Message}", comEx2);
                }
            }

            string tempInput  = Path.Combine(Path.GetTempPath(), $"magosa_in_{Guid.NewGuid():N}.png");
            string tempOutput = Path.Combine(Path.GetTempPath(), $"magosa_out_{Guid.NewGuid():N}.png");
            Log($"tempInput ={tempInput}");
            Log($"tempOutput={tempOutput}");

            try
            {
                // ── STEP 3: shape.Export ─────────────────────────────────────────
                try
                {
                    shape.Export(tempInput, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    long exportedSize = File.Exists(tempInput) ? new FileInfo(tempInput).Length : -1;
                    Log($"STEP3 OK: Export 完了, ファイルサイズ={exportedSize} bytes");
                }
                catch (COMException comEx)
                {
                    Log($"STEP3 FAIL (COMException): HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                    throw new InvalidOperationException(
                        $"[STEP3 shape.Export] HRESULT=0x{comEx.HResult:X8}: {comEx.Message}", comEx);
                }

                // ── STEP 4: Magick.NET 処理 ──────────────────────────────────────
                try
                {
                    ProcessImage(tempInput, tempOutput, settings);
                    long processedSize = File.Exists(tempOutput) ? new FileInfo(tempOutput).Length : -1;
                    Log($"STEP4 OK: ProcessImage 完了, ファイルサイズ={processedSize} bytes");
                }
                catch (Exception ex)
                {
                    Log($"STEP4 FAIL: {ex.GetType().Name} {ex.Message}");
                    throw new InvalidOperationException(
                        $"[STEP4 ProcessImage] {ex.GetType().Name}: {ex.Message}", ex);
                }

                // ── STEP 5: AddPicture ───────────────────────────────────────────
                PowerPoint.Shape newShape;
                try
                {
                    newShape = slide.Shapes.AddPicture(
                        tempOutput,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoCTrue,
                        left, top, width, height);
                    Log($"STEP5 OK: AddPicture 完了, newShape.Name={newShape.Name}");
                }
                catch (COMException comEx)
                {
                    Log($"STEP5 FAIL (COMException): HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                    throw new InvalidOperationException(
                        $"[STEP5 AddPicture] HRESULT=0x{comEx.HResult:X8}: {comEx.Message}", comEx);
                }

                // ── STEP 6: 元図形を削除 ─────────────────────────────────────────
                try
                {
                    shape.Delete();
                    Log("STEP6 OK: shape.Delete() 完了");
                }
                catch (COMException comEx)
                {
                    Log($"STEP6 FAIL (COMException): HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                    throw new InvalidOperationException(
                        $"[STEP6 shape.Delete] HRESULT=0x{comEx.HResult:X8}: {comEx.Message}", comEx);
                }

                // ── STEP 7: 新図形のプロパティ復元 ──────────────────────────────
                try
                {
                    newShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    newShape.Width  = width;
                    newShape.Height = height;
                    newShape.Left   = left;
                    newShape.Top    = top;
                    Log($"STEP7 OK: サイズ復元 W={width:F1} H={height:F1} L={left:F1} T={top:F1}");
                }
                catch (COMException comEx)
                {
                    Log($"STEP7 FAIL (COMException): HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                    throw new InvalidOperationException(
                        $"[STEP7 newShape プロパティ復元] HRESULT=0x{comEx.HResult:X8}: {comEx.Message}", comEx);
                }

                // ── STEP 8: Z 順を復元 ──────────────────────────────────────────
                try
                {
                    int maxIter = 300;
                    int moved = 0;
                    while (newShape.ZOrderPosition > zOrder && maxIter-- > 0)
                    {
                        newShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                        moved++;
                    }
                    Log($"STEP8 OK: Z順復元 moved={moved} 回, ZOrderPosition={newShape.ZOrderPosition}");
                }
                catch (COMException comEx)
                {
                    Log($"STEP8 FAIL (COMException): HRESULT=0x{comEx.HResult:X8} {comEx.Message}");
                    throw new InvalidOperationException(
                        $"[STEP8 Z順復元] HRESULT=0x{comEx.HResult:X8}: {comEx.Message}", comEx);
                }

                Log("=== Apply 正常完了 ===");
            }
            finally
            {
                try { if (File.Exists(tempInput))  File.Delete(tempInput);  } catch { }
                try { if (File.Exists(tempOutput)) File.Delete(tempOutput); } catch { }
            }
        }

        // ─── Magick.NET パイプライン ────────────────────────────────────────────────

        private void ProcessImage(string inputPath, string outputPath, ImageColorSettings settings)
        {
            using (var image = new MagickImage(inputPath))
            {
                Log($"ProcessImage: 入力={image.Width}x{image.Height}, CS={image.ColorSpace}, Depth={image.Depth}, HasAlpha={image.HasAlpha}");

                // アルファチャンネルを白背景に合成してフラット化（透明領域が閾値処理に干渉するのを防ぐ）
                if (image.HasAlpha)
                {
                    image.BackgroundColor = MagickColors.White;
                    image.Alpha(AlphaOption.Remove);
                    Log("ProcessImage: HasAlpha=True → アルファを白背景にフラット化");
                }

                // Step 1: グレースケール
                if (settings.ToneMode == ColorToneMode.Grayscale)
                    image.Grayscale();

                // Step 2: 白黒 or セピア（排他）
                if (settings.ToneMode == ColorToneMode.BlackAndWhite)
                {
                    image.Grayscale();
                    // Normalize でヒストグラムを 0〜65535 全域に拡張してから二値化する。
                    // これにより、画像全体の明度に依存せず閾値が意味を持つようになる。
                    image.Normalize();
                    Log($"ProcessImage Step2: Grayscale+Normalize完了, Threshold={settings.BlackWhiteThreshold}%");
                    image.Threshold(new Percentage(settings.BlackWhiteThreshold));
                }
                else if (settings.ToneMode == ColorToneMode.Sepia)
                    image.SepiaTone();

                // Step 3: 明るさ・コントラスト
                if (settings.Brightness != 0 || settings.Contrast != 0)
                    image.BrightnessContrast(
                        new Percentage(settings.Brightness),
                        new Percentage(settings.Contrast));

                // Step 4: 色相・彩度
                if (settings.Hue != 0 || settings.Saturation != 0)
                    image.Modulate(
                        new Percentage(100),
                        new Percentage(settings.Saturation + 100),
                        new Percentage(settings.Hue + 180));

                // Step 5: カラーライズ（Photoshop「色彩の統一」相当）
                // 各ピクセルの H と S を指定色で統一し、L（輝度）は元画像を保持する。
                // ImageMagick の Colorize() は RGB ブレンドなので Photoshop と挙動が異なるため、
                // HLS カラースペース上でチャンネルを直接置換する方式を採用する。
                if (settings.ColorizeEnabled)
                {
                    int r = settings.ColorizeRgb & 0xFF;
                    int g = (settings.ColorizeRgb >> 8)  & 0xFF;
                    int b = (settings.ColorizeRgb >> 16) & 0xFF;

                    // RGB → Hue 計算（0〜360°）
                    double rf = r / 255.0, gf = g / 255.0, bf = b / 255.0;
                    double cmax = Math.Max(Math.Max(rf, gf), bf);
                    double cmin = Math.Min(Math.Min(rf, gf), bf);
                    double delta = cmax - cmin;

                    double targetH = 0.0;
                    if (delta > 1e-9)
                    {
                        if      (cmax == rf) targetH = 60.0 * ((gf - bf) / delta);
                        else if (cmax == gf) targetH = 60.0 * ((bf - rf) / delta + 2.0);
                        else                 targetH = 60.0 * ((rf - gf) / delta + 4.0);
                        if (targetH < 0.0) targetH += 360.0;
                    }

                    // 強度（0〜100）を彩度（0〜1）にマッピング
                    // Photoshop のカラーライズは「彩度スライダー」が強度に対応する
                    double targetS = settings.ColorizeIntensity / 100.0;

                    Log($"ProcessImage カラーライズ: RGB=({r},{g},{b}) → H={targetH:F1}° S={targetS:F2}");

                    // ImageMagick HSL カラースペースのチャンネル配置:
                    //   Red   = H (Hue,        0-360° → Percentage 0-100%)
                    //   Green = S (Saturation,  0-1   → Percentage 0-100%)
                    //   Blue  = L (Lightness,   変更しない)
                    // H と S を定数で置換し、L（Blue）は一切変更しない。
                    image.ColorSpace = ColorSpace.HSL;
                    image.Evaluate(Channels.Red,   EvaluateOperator.Set, new Percentage(targetH / 360.0 * 100.0));
                    image.Evaluate(Channels.Green, EvaluateOperator.Set, new Percentage(targetS * 100.0));
                    image.ColorSpace = ColorSpace.sRGB;
                }

                image.Write(outputPath);
                Log($"ProcessImage: 出力={image.Width}x{image.Height}, ファイル書き込み完了");
            }
        }
    }
}
