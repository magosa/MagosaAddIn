using System;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 色空間変換ユーティリティクラス
    /// PowerPointのRGB形式（int値）とHSL色空間の相互変換を提供
    /// </summary>
    public static class ColorConverter
    {
        #region RGB ⇔ HSL 変換

        /// <summary>
        /// PowerPoint RGB値（int）をHSL色空間に変換
        /// </summary>
        /// <param name="rgb">PowerPoint RGB値（0xBBGGRR形式）</param>
        /// <returns>(H: 0-360, S: 0-1, L: 0-1)</returns>
        public static (float H, float S, float L) RgbToHsl(int rgb)
        {
            // PowerPointのRGB形式は 0xBBGGRR（BGRバイト順）
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;

            return RgbToHsl(r, g, b);
        }

        /// <summary>
        /// RGB値（0-255）をHSL色空間に変換
        /// </summary>
        /// <param name="r">赤（0-255）</param>
        /// <param name="g">緑（0-255）</param>
        /// <param name="b">青（0-255）</param>
        /// <returns>(H: 0-360, S: 0-1, L: 0-1)</returns>
        public static (float H, float S, float L) RgbToHsl(int r, int g, int b)
        {
            float rf = r / 255f;
            float gf = g / 255f;
            float bf = b / 255f;

            float max = Math.Max(rf, Math.Max(gf, bf));
            float min = Math.Min(rf, Math.Min(gf, bf));
            float delta = max - min;

            // 明度（Lightness）
            float l = (max + min) / 2f;

            // 彩度（Saturation）
            float s = 0f;
            if (delta != 0f)
            {
                s = l > 0.5f ? delta / (2f - max - min) : delta / (max + min);
            }

            // 色相（Hue）
            float h = 0f;
            if (delta != 0f)
            {
                if (max == rf)
                {
                    h = ((gf - bf) / delta) + (gf < bf ? 6f : 0f);
                }
                else if (max == gf)
                {
                    h = ((bf - rf) / delta) + 2f;
                }
                else
                {
                    h = ((rf - gf) / delta) + 4f;
                }
                h *= 60f;
            }

            return (h, s, l);
        }

        /// <summary>
        /// HSL色空間をPowerPoint RGB値（int）に変換
        /// </summary>
        /// <param name="h">色相（0-360）</param>
        /// <param name="s">彩度（0-1）</param>
        /// <param name="l">明度（0-1）</param>
        /// <returns>PowerPoint RGB値（0xBBGGRR形式）</returns>
        public static int HslToRgb(float h, float s, float l)
        {
            // 範囲制限
            h = h % 360f;
            if (h < 0) h += 360f;
            s = Math.Max(0f, Math.Min(1f, s));
            l = Math.Max(0f, Math.Min(1f, l));

            float c = (1f - Math.Abs(2f * l - 1f)) * s;
            float x = c * (1f - Math.Abs((h / 60f) % 2f - 1f));
            float m = l - c / 2f;

            float rf, gf, bf;

            if (h < 60f)
            {
                rf = c; gf = x; bf = 0f;
            }
            else if (h < 120f)
            {
                rf = x; gf = c; bf = 0f;
            }
            else if (h < 180f)
            {
                rf = 0f; gf = c; bf = x;
            }
            else if (h < 240f)
            {
                rf = 0f; gf = x; bf = c;
            }
            else if (h < 300f)
            {
                rf = x; gf = 0f; bf = c;
            }
            else
            {
                rf = c; gf = 0f; bf = x;
            }

            int r = (int)Math.Round((rf + m) * 255f);
            int g = (int)Math.Round((gf + m) * 255f);
            int b = (int)Math.Round((bf + m) * 255f);

            // PowerPoint形式に変換（0xBBGGRR）
            return r | (g << 8) | (b << 16);
        }

        #endregion

        #region RGB ⇔ HSV 変換

        /// <summary>
        /// PowerPoint RGB値（int）をHSV色空間に変換
        /// </summary>
        /// <param name="rgb">PowerPoint RGB値（0xBBGGRR形式）</param>
        /// <returns>(H: 0-360, S: 0-1, V: 0-1)</returns>
        public static (float H, float S, float V) RgbToHsv(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;

            return RgbToHsv(r, g, b);
        }

        /// <summary>
        /// RGB値（0-255）をHSV色空間に変換
        /// </summary>
        public static (float H, float S, float V) RgbToHsv(int r, int g, int b)
        {
            float rf = r / 255f;
            float gf = g / 255f;
            float bf = b / 255f;

            float max = Math.Max(rf, Math.Max(gf, bf));
            float min = Math.Min(rf, Math.Min(gf, bf));
            float delta = max - min;

            // 明度（Value）
            float v = max;

            // 彩度（Saturation）
            float s = (max != 0f) ? (delta / max) : 0f;

            // 色相（Hue）
            float h = 0f;
            if (delta != 0f)
            {
                if (max == rf)
                {
                    h = ((gf - bf) / delta) + (gf < bf ? 6f : 0f);
                }
                else if (max == gf)
                {
                    h = ((bf - rf) / delta) + 2f;
                }
                else
                {
                    h = ((rf - gf) / delta) + 4f;
                }
                h *= 60f;
            }

            return (h, s, v);
        }

        /// <summary>
        /// HSV色空間をPowerPoint RGB値（int）に変換
        /// </summary>
        public static int HsvToRgb(float h, float s, float v)
        {
            h = h % 360f;
            if (h < 0) h += 360f;
            s = Math.Max(0f, Math.Min(1f, s));
            v = Math.Max(0f, Math.Min(1f, v));

            float c = v * s;
            float x = c * (1f - Math.Abs((h / 60f) % 2f - 1f));
            float m = v - c;

            float rf, gf, bf;

            if (h < 60f)
            {
                rf = c; gf = x; bf = 0f;
            }
            else if (h < 120f)
            {
                rf = x; gf = c; bf = 0f;
            }
            else if (h < 180f)
            {
                rf = 0f; gf = c; bf = x;
            }
            else if (h < 240f)
            {
                rf = 0f; gf = x; bf = c;
            }
            else if (h < 300f)
            {
                rf = x; gf = 0f; bf = c;
            }
            else
            {
                rf = c; gf = 0f; bf = x;
            }

            int r = (int)Math.Round((rf + m) * 255f);
            int g = (int)Math.Round((gf + m) * 255f);
            int b = (int)Math.Round((bf + m) * 255f);

            return r | (g << 8) | (b << 16);
        }

        #endregion

        #region ユーティリティメソッド

        /// <summary>
        /// RGB値を16進数文字列に変換（#RRGGBB形式）
        /// </summary>
        public static string RgbToHex(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>
        /// 16進数文字列をRGB値に変換
        /// </summary>
        /// <param name="hex">#RRGGBB または RRGGBB 形式</param>
        public static int HexToRgb(string hex)
        {
            hex = hex.TrimStart('#');
            if (hex.Length != 6)
                throw new ArgumentException("16進数カラーコードは6桁で指定してください（例: #FF5733）");

            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);

            return r | (g << 8) | (b << 16);
        }

        /// <summary>
        /// System.Drawing.Color をPowerPoint RGB値に変換
        /// </summary>
        public static int ColorToRgb(System.Drawing.Color color)
        {
            return color.R | (color.G << 8) | (color.B << 16);
        }

        /// <summary>
        /// PowerPoint RGB値をSystem.Drawing.Colorに変換
        /// </summary>
        public static System.Drawing.Color RgbToColor(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return System.Drawing.Color.FromArgb(r, g, b);
        }

        /// <summary>
        /// 補色を計算（色相を180度回転）
        /// </summary>
        public static int GetComplementary(int rgb)
        {
            var (h, s, l) = RgbToHsl(rgb);
            return HslToRgb((h + 180f) % 360f, s, l);
        }

        #endregion
    }
}
