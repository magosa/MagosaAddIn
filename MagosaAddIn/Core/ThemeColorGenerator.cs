using System;
using System.Collections.Generic;
using System.Linq;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// テーマカラー生成クラス
    /// 17種類の配色パターンに基づいてテーマカラーを生成
    /// </summary>
    public class ThemeColorGenerator
    {
        #region メイン生成メソッド

        /// <summary>
        /// 指定された配色パターンでカラーテーマを生成
        /// </summary>
        /// <param name="baseColor">ベースカラー（PowerPoint RGB値）</param>
        /// <param name="schemeType">配色パターン</param>
        /// <param name="colorCount">生成する色数（デフォルト5色）</param>
        /// <returns>生成されたカラーリスト</returns>
        public List<int> GenerateColorScheme(int baseColor, ColorSchemeType schemeType, int colorCount = 5)
        {
            switch (schemeType)
            {
                // 色相ベース配色
                case ColorSchemeType.Dyad:
                    return GenerateDyad(baseColor);
                case ColorSchemeType.Triad:
                    return GenerateTriad(baseColor);
                case ColorSchemeType.Tetrad:
                    return GenerateTetrad(baseColor);
                case ColorSchemeType.Pentad:
                    return GeneratePentad(baseColor);
                case ColorSchemeType.Hexad:
                    return GenerateHexad(baseColor);
                case ColorSchemeType.Analogy:
                    return GenerateAnalogy(baseColor, colorCount);
                case ColorSchemeType.Intermediate:
                    return GenerateIntermediate(baseColor);
                case ColorSchemeType.Opponent:
                    return GenerateOpponent(baseColor);
                case ColorSchemeType.SplitComplementary:
                    return GenerateSplitComplementary(baseColor);

                // トーンベース配色
                case ColorSchemeType.ToneOnTone:
                    return GenerateToneOnTone(baseColor, colorCount);
                case ColorSchemeType.ToneInTone:
                    return GenerateToneInTone(baseColor, colorCount);
                case ColorSchemeType.Camaieu:
                    return GenerateCamaieu(baseColor, colorCount);
                case ColorSchemeType.FauxCamaieu:
                    return GenerateFauxCamaieu(baseColor, colorCount);
                case ColorSchemeType.DominantColor:
                    return GenerateDominantColor(baseColor, colorCount);
                case ColorSchemeType.Identity:
                    return GenerateIdentity(baseColor, colorCount);
                case ColorSchemeType.Gradation:
                    return GenerateGradation(baseColor, colorCount);

                // コントラスト配色
                case ColorSchemeType.HueContrast:
                    return GenerateHueContrast(baseColor);
                case ColorSchemeType.LightnessContrast:
                    return GenerateLightnessContrast(baseColor);
                case ColorSchemeType.SaturationContrast:
                    return GenerateSaturationContrast(baseColor);

                default:
                    throw new ArgumentException($"未対応の配色パターン: {schemeType}");
            }
        }

        #endregion

        #region 色相ベース配色

        /// <summary>
        /// ダイアード（補色）- 色相環で180°反対
        /// </summary>
        private List<int> GenerateDyad(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 180f) % 360f, s, l)
            };
        }

        /// <summary>
        /// トライアド（3等分）- 色相環で120°間隔
        /// </summary>
        private List<int> GenerateTriad(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 120f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 240f) % 360f, s, l)
            };
        }

        /// <summary>
        /// テトラード（4等分）- 色相環で90°間隔
        /// </summary>
        private List<int> GenerateTetrad(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 90f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 180f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 270f) % 360f, s, l)
            };
        }

        /// <summary>
        /// ペンタード（5等分）- 色相環で72°間隔
        /// </summary>
        private List<int> GeneratePentad(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 72f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 144f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 216f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 288f) % 360f, s, l)
            };
        }

        /// <summary>
        /// ヘクサード（6等分）- 色相環で60°間隔
        /// </summary>
        private List<int> GenerateHexad(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 60f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 120f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 180f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 240f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 300f) % 360f, s, l)
            };
        }

        /// <summary>
        /// アナロジー（類似色）- 色相環で30°間隔
        /// </summary>
        private List<int> GenerateAnalogy(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int>();
            
            int halfCount = count / 2;
            for (int i = -halfCount; i <= count - halfCount - 1; i++)
            {
                float newH = (h + (i * 30f)) % 360f;
                if (newH < 0) newH += 360f;
                colors.Add(ColorConverter.HslToRgb(newH, s, l));
            }
            
            return colors;
        }

        /// <summary>
        /// インターミディエート - 色相環で90°
        /// </summary>
        private List<int> GenerateIntermediate(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 90f) % 360f, s, l)
            };
        }

        /// <summary>
        /// オポーネント - 色相環で135°
        /// </summary>
        private List<int> GenerateOpponent(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 135f) % 360f, s, l)
            };
        }

        /// <summary>
        /// スプリットコンプリメンタリー - 補色の両隣（180°±30°）
        /// </summary>
        private List<int> GenerateSplitComplementary(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb((h + 150f) % 360f, s, l),
                ColorConverter.HslToRgb((h + 210f) % 360f, s, l)
            };
        }

        #endregion

        #region トーンベース配色

        /// <summary>
        /// トーンオントーン - 同色相・明度差大
        /// </summary>
        private List<int> GenerateToneOnTone(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int>();
            
            // 明度を20%～90%の範囲で分散
            float minLightness = 0.2f;
            float maxLightness = 0.9f;
            float step = (maxLightness - minLightness) / (count - 1);
            
            for (int i = 0; i < count; i++)
            {
                float newL = minLightness + (step * i);
                colors.Add(ColorConverter.HslToRgb(h, s, newL));
            }
            
            return colors;
        }

        /// <summary>
        /// トーンイントーン - トーン統一・色相変化
        /// </summary>
        private List<int> GenerateToneInTone(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int>();
            
            // 色相を均等に分散、明度・彩度は固定
            float hueStep = 360f / count;
            
            for (int i = 0; i < count; i++)
            {
                float newH = (h + (hueStep * i)) % 360f;
                colors.Add(ColorConverter.HslToRgb(newH, s, l));
            }
            
            return colors;
        }

        /// <summary>
        /// カマイユ - 色相・トーン近似（色相差±5°、彩度差±10%）
        /// </summary>
        private List<int> GenerateCamaieu(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int> { baseColor };
            var random = new Random();
            
            for (int i = 1; i < count; i++)
            {
                float hueVariation = (float)((random.NextDouble() - 0.5) * 10); // ±5°
                float satVariation = (float)((random.NextDouble() - 0.5) * 0.2); // ±10%
                float lightVariation = (float)((random.NextDouble() - 0.5) * 0.2); // ±10%
                
                float newH = (h + hueVariation) % 360f;
                if (newH < 0) newH += 360f;
                float newS = Math.Max(0f, Math.Min(1f, s + satVariation));
                float newL = Math.Max(0f, Math.Min(1f, l + lightVariation));
                
                colors.Add(ColorConverter.HslToRgb(newH, newS, newL));
            }
            
            return colors;
        }

        /// <summary>
        /// フォカマイユ - カマイユより色相変化大（色相差±15°、彩度差±20%）
        /// </summary>
        private List<int> GenerateFauxCamaieu(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int> { baseColor };
            var random = new Random();
            
            for (int i = 1; i < count; i++)
            {
                float hueVariation = (float)((random.NextDouble() - 0.5) * 30); // ±15°
                float satVariation = (float)((random.NextDouble() - 0.5) * 0.4); // ±20%
                float lightVariation = (float)((random.NextDouble() - 0.5) * 0.3); // ±15%
                
                float newH = (h + hueVariation) % 360f;
                if (newH < 0) newH += 360f;
                float newS = Math.Max(0f, Math.Min(1f, s + satVariation));
                float newL = Math.Max(0f, Math.Min(1f, l + lightVariation));
                
                colors.Add(ColorConverter.HslToRgb(newH, newS, newL));
            }
            
            return colors;
        }

        /// <summary>
        /// ドミナントカラー - 同色相統一・彩度と明度を変化
        /// </summary>
        private List<int> GenerateDominantColor(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int> { baseColor };
            
            for (int i = 1; i < count; i++)
            {
                float satRatio = 0.3f + (0.7f * i / (count - 1)); // 30%～100%
                float lightRatio = 0.3f + (0.6f * i / (count - 1)); // 30%～90%
                
                colors.Add(ColorConverter.HslToRgb(h, s * satRatio, lightRatio));
            }
            
            return colors;
        }

        /// <summary>
        /// アイデンティティ - 1色相・明度彩度変化
        /// </summary>
        private List<int> GenerateIdentity(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int>();
            
            // 明度と彩度を組み合わせて変化
            for (int i = 0; i < count; i++)
            {
                float ratio = i / (float)(count - 1);
                float newS = 0.3f + (s * 0.7f * ratio);
                float newL = 0.2f + (0.7f * ratio);
                
                colors.Add(ColorConverter.HslToRgb(h, newS, newL));
            }
            
            return colors;
        }

        /// <summary>
        /// グラデーション - 色相・明度・彩度を段階的に変化
        /// </summary>
        private List<int> GenerateGradation(int baseColor, int count)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            var colors = new List<int>();
            
            // 終点の色（色相を180°回転）
            float endH = (h + 180f) % 360f;
            float endS = Math.Max(0.2f, s * 0.7f);
            float endL = 1f - l; // 明度を反転
            
            for (int i = 0; i < count; i++)
            {
                float ratio = i / (float)(count - 1);
                float newH = h + ((endH - h) * ratio);
                if (newH < 0) newH += 360f;
                newH = newH % 360f;
                float newS = s + ((endS - s) * ratio);
                float newL = l + ((endL - l) * ratio);
                
                colors.Add(ColorConverter.HslToRgb(newH, newS, newL));
            }
            
            return colors;
        }

        #endregion

        #region コントラスト配色

        /// <summary>
        /// 色相コントラスト - 補色関係
        /// </summary>
        private List<int> GenerateHueContrast(int baseColor)
        {
            return GenerateDyad(baseColor);
        }

        /// <summary>
        /// 明度コントラスト - 明暗対比
        /// </summary>
        private List<int> GenerateLightnessContrast(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            
            // 明度を反転（明るい色→暗い色、暗い色→明るい色）
            float contrastL = l > 0.5f ? 0.2f : 0.8f;
            
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb(h, s, contrastL)
            };
        }

        /// <summary>
        /// 彩度コントラスト - 鮮やか⇔くすんだ
        /// </summary>
        private List<int> GenerateSaturationContrast(int baseColor)
        {
            var (h, s, l) = ColorConverter.RgbToHsl(baseColor);
            
            // 彩度を反転（鮮やか→くすみ、くすみ→鮮やか）
            float contrastS = s > 0.5f ? 0.2f : 0.9f;
            
            return new List<int>
            {
                baseColor,
                ColorConverter.HslToRgb(h, contrastS, l)
            };
        }

        #endregion

        #region 明度バリエーション生成

        /// <summary>
        /// 明度バリエーションを生成（基準色を中心に明暗方向に展開）
        /// </summary>
        /// <param name="baseColors">ベースカラーのリスト</param>
        /// <param name="steps">明度段階数（1～10段階）</param>
        /// <returns>色×明度段階の2次元リスト</returns>
        public List<List<int>> GenerateLightnessVariations(List<int> baseColors, int steps = 3)
        {
            var result = new List<List<int>>();
            
            foreach (var color in baseColors)
            {
                var variations = new List<int>();
                var (h, s, l) = ColorConverter.RgbToHsl(color);
                
                if (steps == 1)
                {
                    // 段階数が1の場合は基準色のみ
                    variations.Add(color);
                }
                else if (steps == 2)
                {
                    // 段階数が2の場合は基準色と暗い色
                    variations.Add(color);
                    float darkerL = Math.Max(0.1f, l * 0.6f);
                    variations.Add(ColorConverter.HslToRgb(h, s, darkerL));
                }
                else
                {
                    // 段階数が3以上の場合は基準色を中心に明暗両方向に展開
                    // 明るい側と暗い側の段階数を計算
                    int lighterSteps = steps / 2;
                    int darkerSteps = steps - lighterSteps - 1; // 基準色を除く
                    
                    // 明るい方向に展開
                    for (int i = lighterSteps; i > 0; i--)
                    {
                        float ratio = i / (float)(lighterSteps + 1);
                        float newL = l + (0.95f - l) * ratio;
                        newL = Math.Min(0.95f, newL);
                        variations.Add(ColorConverter.HslToRgb(h, s, newL));
                    }
                    
                    // 基準色
                    variations.Add(color);
                    
                    // 暗い方向に展開
                    for (int i = 1; i <= darkerSteps; i++)
                    {
                        float ratio = i / (float)(darkerSteps + 1);
                        float newL = l - (l - 0.1f) * ratio;
                        newL = Math.Max(0.1f, newL);
                        variations.Add(ColorConverter.HslToRgb(h, s, newL));
                    }
                }
                
                result.Add(variations);
            }
            
            return result;
        }

        #endregion

        #region ヘルパーメソッド

        /// <summary>
        /// 配色パターンの表示名を取得
        /// </summary>
        public static string GetSchemeDisplayName(ColorSchemeType schemeType)
        {
            switch (schemeType)
            {
                case ColorSchemeType.Dyad:
                    return "ダイアード（補色）";
                case ColorSchemeType.Triad:
                    return "トライアド（3等分）";
                case ColorSchemeType.Tetrad:
                    return "テトラード（4等分）";
                case ColorSchemeType.Pentad:
                    return "ペンタード（5等分）";
                case ColorSchemeType.Hexad:
                    return "ヘクサード（6等分）";
                case ColorSchemeType.Analogy:
                    return "アナロジー（類似色）";
                case ColorSchemeType.Intermediate:
                    return "インターミディエート（90°）";
                case ColorSchemeType.Opponent:
                    return "オポーネント（135°）";
                case ColorSchemeType.SplitComplementary:
                    return "スプリットコンプリメンタリー";
                case ColorSchemeType.ToneOnTone:
                    return "トーンオントーン";
                case ColorSchemeType.ToneInTone:
                    return "トーンイントーン";
                case ColorSchemeType.Camaieu:
                    return "カマイユ";
                case ColorSchemeType.FauxCamaieu:
                    return "フォカマイユ";
                case ColorSchemeType.DominantColor:
                    return "ドミナントカラー";
                case ColorSchemeType.Identity:
                    return "アイデンティティ";
                case ColorSchemeType.Gradation:
                    return "グラデーション";
                case ColorSchemeType.HueContrast:
                    return "色相コントラスト";
                case ColorSchemeType.LightnessContrast:
                    return "明度コントラスト";
                case ColorSchemeType.SaturationContrast:
                    return "彩度コントラスト";
                default:
                    return schemeType.ToString();
            }
        }

        #endregion
    }
}
