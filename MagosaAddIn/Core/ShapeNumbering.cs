using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形への自動ナンバリング機能を提供するクラス
    /// </summary>
    public class ShapeNumbering
    {
        /// <summary>
        /// 図形に自動ナンバリングを適用
        /// </summary>
        /// <param name="shapes">ナンバリング対象の図形リスト</param>
        /// <param name="startNumber">開始番号</param>
        /// <param name="increment">増分値</param>
        /// <param name="format">番号フォーマット</param>
        /// <param name="fontSize">フォントサイズ（pt）</param>
        public void ApplyNumbering(List<PowerPoint.Shape> shapes, int startNumber, int increment, 
            NumberFormat format, float fontSize)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_NUMBERING, "自動ナンバリング");
                    ErrorHandler.ValidateRange(startNumber, Constants.MIN_START_NUMBER, 
                        Constants.MAX_START_NUMBER, "開始番号", "自動ナンバリング");
                    ErrorHandler.ValidateRange(increment, Constants.MIN_INCREMENT, 
                        Constants.MAX_INCREMENT, "増分値", "自動ナンバリング");

                    int currentNumber = startNumber;

                    for (int i = 0; i < shapes.Count; i++)
                    {
                        var shape = shapes[i];
                        string numberText = FormatNumber(currentNumber, format);

                        // 図形にテキストを設定
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            var textFrame = shape.TextFrame;
                            if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                // 既存テキストがある場合は先頭に追加
                                textFrame.TextRange.Text = numberText + "\n" + textFrame.TextRange.Text;
                            }
                            else
                            {
                                // テキストがない場合は新規設定
                                textFrame.TextRange.Text = numberText;
                            }

                            // フォントサイズを設定
                            textFrame.TextRange.Font.Size = fontSize;

                            // 中央揃えに設定
                            textFrame.TextRange.ParagraphFormat.Alignment = 
                                PowerPoint.PpParagraphAlignment.ppAlignCenter;
                            textFrame.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                        }

                        currentNumber += increment;
                    }

                    ComExceptionHandler.LogDebug($"自動ナンバリング完了: {shapes.Count}個, " +
                        $"開始={startNumber}, 増分={increment}, フォーマット={format}");
                },
                "自動ナンバリング");
        }

        /// <summary>
        /// 番号を指定フォーマットで文字列に変換
        /// </summary>
        private string FormatNumber(int number, NumberFormat format)
        {
            switch (format)
            {
                case NumberFormat.Arabic:
                    return number.ToString();

                case NumberFormat.CircledArabic:
                    return GetCircledNumber(number);

                case NumberFormat.UpperAlpha:
                    return GetAlphabeticNumber(number, true);

                case NumberFormat.LowerAlpha:
                    return GetAlphabeticNumber(number, false);

                case NumberFormat.UpperRoman:
                    return GetRomanNumber(number, true);

                case NumberFormat.LowerRoman:
                    return GetRomanNumber(number, false);

                default:
                    return number.ToString();
            }
        }

        /// <summary>
        /// 丸数字を取得（①②③...）
        /// </summary>
        private string GetCircledNumber(int number)
        {
            // Unicode丸数字は①(U+2460)～⑳(U+2473)まで
            if (number >= 1 && number <= 20)
            {
                return char.ConvertFromUtf32(0x245F + number);
            }
            else if (number >= 21 && number <= 35)
            {
                // ㉑～㉟ (U+3251～U+325F)
                return char.ConvertFromUtf32(0x3250 + (number - 20));
            }
            else if (number >= 36 && number <= 50)
            {
                // ㊱～㊿ (U+32B1～U+32BF)
                return char.ConvertFromUtf32(0x32B0 + (number - 35));
            }
            else
            {
                // 範囲外の場合は括弧付き数字で代用
                return $"({number})";
            }
        }

        /// <summary>
        /// アルファベット番号を取得（A, B, C...またはa, b, c...）
        /// </summary>
        private string GetAlphabeticNumber(int number, bool isUpper)
        {
            if (number < 1)
                return isUpper ? "A" : "a";

            string result = "";
            int n = number;

            while (n > 0)
            {
                n--;  // 0ベースに調整
                int remainder = n % 26;
                char c = (char)((isUpper ? 'A' : 'a') + remainder);
                result = c + result;
                n = n / 26;
            }

            return result;
        }

        /// <summary>
        /// ローマ数字を取得（I, II, III...またはi, ii, iii...）
        /// </summary>
        private string GetRomanNumber(int number, bool isUpper)
        {
            if (number < 1 || number > 3999)
                return number.ToString();  // 範囲外は算用数字

            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] romanUpper = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            string[] romanLower = { "m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i" };

            string[] romans = isUpper ? romanUpper : romanLower;
            string result = "";

            for (int i = 0; i < values.Length; i++)
            {
                while (number >= values[i])
                {
                    number -= values[i];
                    result += romans[i];
                }
            }

            return result;
        }

        /// <summary>
        /// 図形からテキストを削除
        /// </summary>
        public void ClearNumbering(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_NUMBERING, "ナンバリング削除");

                    foreach (var shape in shapes)
                    {
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Text = "";
                        }
                    }

                    ComExceptionHandler.LogDebug($"ナンバリング削除完了: {shapes.Count}個");
                },
                "ナンバリング削除");
        }
    }
}