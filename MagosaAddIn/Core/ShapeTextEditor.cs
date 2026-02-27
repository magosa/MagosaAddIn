using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形テキストの一括編集機能を提供するクラス
    /// </summary>
    public class ShapeTextEditor
    {
        #region テキスト情報取得

        /// <summary>
        /// 図形リストからテキスト情報を取得
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <returns>テキスト情報リスト</returns>
        public List<ShapeTextInfo> GetTextInfos(List<PowerPoint.Shape> shapes)
        {
            var result = new List<ShapeTextInfo>();

            foreach (var shape in shapes)
            {
                var info = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        bool hasTextFrame = shape.HasTextFrame == Office.MsoTriState.msoTrue;
                        string text = "";
                        bool hasText = false;

                        if (hasTextFrame)
                        {
                            hasText = shape.TextFrame.HasText == Office.MsoTriState.msoTrue;
                            if (hasText)
                            {
                                text = shape.TextFrame.TextRange.Text;
                            }
                        }

                        return new ShapeTextInfo
                        {
                            Shape = shape,
                            ShapeName = shape.Name,
                            Text = text,
                            HasTextFrame = hasTextFrame,
                            HasText = hasText
                        };
                    },
                    $"テキスト情報取得: {shape.Name}",
                    defaultValue: new ShapeTextInfo { ShapeName = shape.Name, HasTextFrame = false },
                    suppressErrors: true);

                result.Add(info);
            }

            ComExceptionHandler.LogDebug($"テキスト情報取得完了: {result.Count}個");
            return result;
        }

        #endregion

        #region テキスト設定

        /// <summary>
        /// 各図形に個別テキストを適用
        /// </summary>
        /// <param name="textInfos">テキスト情報リスト（Textプロパティが更新済み）</param>
        /// <returns>成功した図形数</returns>
        public int ApplyIndividualTexts(List<ShapeTextInfo> textInfos)
        {
            int successCount = 0;

            foreach (var info in textInfos)
            {
                if (!info.HasTextFrame) continue;

                var success = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        info.Shape.TextFrame.TextRange.Text = info.Text ?? "";
                        return true;
                    },
                    $"テキスト適用: {info.ShapeName}",
                    defaultValue: false,
                    suppressErrors: true);

                if (success) successCount++;
            }

            ComExceptionHandler.LogDebug($"個別テキスト適用完了: {successCount}/{textInfos.Count}個");
            return successCount;
        }

        /// <summary>
        /// 全図形に同一テキストを設定
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="text">設定するテキスト</param>
        /// <returns>成功した図形数</returns>
        public int SetUniformText(List<PowerPoint.Shape> shapes, string text)
        {
            int successCount = 0;

            foreach (var shape in shapes)
            {
                var success = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Text = text ?? "";
                            return true;
                        }
                        return false;
                    },
                    $"テキスト統一: {shape.Name}",
                    defaultValue: false,
                    suppressErrors: true);

                if (success) successCount++;
            }

            ComExceptionHandler.LogDebug($"テキスト統一完了: {successCount}/{shapes.Count}個");
            return successCount;
        }

        /// <summary>
        /// 1つのテキストを複数図形に配布（改行で区切って各図形に割り当て）
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="sourceText">配布元テキスト（改行区切り）</param>
        /// <returns>成功した図形数</returns>
        public int DistributeText(List<PowerPoint.Shape> shapes, string sourceText)
        {
            if (string.IsNullOrEmpty(sourceText)) return 0;

            // 改行で分割
            var lines = sourceText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            int successCount = 0;

            for (int i = 0; i < shapes.Count; i++)
            {
                string textToApply = i < lines.Length ? lines[i] : "";

                var success = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shapes[i].TextFrame.TextRange.Text = textToApply;
                            return true;
                        }
                        return false;
                    },
                    $"テキスト配布: {shapes[i].Name}",
                    defaultValue: false,
                    suppressErrors: true);

                if (success) successCount++;
            }

            ComExceptionHandler.LogDebug($"テキスト配布完了: {successCount}/{shapes.Count}個, ソース行数: {lines.Length}");
            return successCount;
        }

        /// <summary>
        /// 選択図形のテキストを一括削除
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <returns>成功した図形数</returns>
        public int ClearText(List<PowerPoint.Shape> shapes)
        {
            int successCount = 0;

            foreach (var shape in shapes)
            {
                var success = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Text = "";
                            return true;
                        }
                        return false;
                    },
                    $"テキスト削除: {shape.Name}",
                    defaultValue: false,
                    suppressErrors: true);

                if (success) successCount++;
            }

            ComExceptionHandler.LogDebug($"テキスト一括削除完了: {successCount}/{shapes.Count}個");
            return successCount;
        }

        #endregion

        #region 検索・置換

        /// <summary>
        /// 図形内テキストを検索・置換
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="searchText">検索文字列</param>
        /// <param name="replaceText">置換文字列</param>
        /// <param name="caseSensitive">大文字小文字を区別するか</param>
        /// <returns>置換した件数</returns>
        public int SearchAndReplace(List<PowerPoint.Shape> shapes, string searchText, string replaceText, bool caseSensitive)
        {
            if (string.IsNullOrEmpty(searchText)) return 0;

            int replaceCount = 0;
            StringComparison comparison = caseSensitive
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            foreach (var shape in shapes)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return;
                        if (shape.TextFrame.HasText != Office.MsoTriState.msoTrue) return;

                        string currentText = shape.TextFrame.TextRange.Text;
                        if (currentText.IndexOf(searchText, comparison) >= 0)
                        {
                            string newText = caseSensitive
                                ? currentText.Replace(searchText, replaceText)
                                : ReplaceIgnoreCase(currentText, searchText, replaceText);

                            shape.TextFrame.TextRange.Text = newText;
                            replaceCount++;
                        }
                    },
                    $"検索置換: {shape.Name}",
                    suppressErrors: true);
            }

            ComExceptionHandler.LogDebug($"検索置換完了: {replaceCount}個の図形で置換, 検索: '{searchText}' → '{replaceText}'");
            return replaceCount;
        }

        /// <summary>
        /// 図形内テキストを検索（一致する図形のリストを返す）
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="searchText">検索文字列</param>
        /// <param name="caseSensitive">大文字小文字を区別するか</param>
        /// <returns>一致した図形のリスト</returns>
        public List<PowerPoint.Shape> FindShapesByText(List<PowerPoint.Shape> shapes, string searchText, bool caseSensitive)
        {
            if (string.IsNullOrEmpty(searchText)) return new List<PowerPoint.Shape>();

            var result = new List<PowerPoint.Shape>();
            StringComparison comparison = caseSensitive
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            foreach (var shape in shapes)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return;
                        if (shape.TextFrame.HasText != Office.MsoTriState.msoTrue) return;

                        string text = shape.TextFrame.TextRange.Text;
                        if (text.IndexOf(searchText, comparison) >= 0)
                        {
                            result.Add(shape);
                        }
                    },
                    $"テキスト検索: {shape.Name}",
                    suppressErrors: true);
            }

            return result;
        }

        #endregion

        #region フォント書式一括変更

        /// <summary>
        /// フォント設定を一括適用
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="settings">フォント設定（nullのプロパティは変更しない）</param>
        /// <returns>成功した図形数</returns>
        public int ApplyFontSettings(List<PowerPoint.Shape> shapes, FontSettings settings)
        {
            if (settings == null) return 0;

            int successCount = 0;

            foreach (var shape in shapes)
            {
                var success = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return false;

                        var textRange = shape.TextFrame.TextRange;

                        if (!string.IsNullOrEmpty(settings.FontName))
                            textRange.Font.Name = settings.FontName;

                        if (settings.FontSize.HasValue)
                            textRange.Font.Size = settings.FontSize.Value;

                        if (settings.IsBold.HasValue)
                            textRange.Font.Bold = settings.IsBold.Value
                                ? Office.MsoTriState.msoTrue
                                : Office.MsoTriState.msoFalse;

                        if (settings.IsItalic.HasValue)
                            textRange.Font.Italic = settings.IsItalic.Value
                                ? Office.MsoTriState.msoTrue
                                : Office.MsoTriState.msoFalse;

                        if (settings.IsUnderline.HasValue)
                            textRange.Font.Underline = settings.IsUnderline.Value
                                ? Office.MsoTriState.msoTrue
                                : Office.MsoTriState.msoFalse;

                        if (settings.FontColor.HasValue)
                            textRange.Font.Color.RGB = settings.FontColor.Value;

                        return true;
                    },
                    $"フォント設定適用: {shape.Name}",
                    defaultValue: false,
                    suppressErrors: true);

                if (success) successCount++;
            }

            ComExceptionHandler.LogDebug($"フォント設定適用完了: {successCount}/{shapes.Count}個");
            return successCount;
        }

        #endregion

        #region テキストレイアウト設定

        /// <summary>
        /// テキストレイアウト設定を一括適用（行間・余白）
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="settings">レイアウト設定（nullのプロパティは変更しない）</param>
        /// <returns>成功した図形数</returns>
        public int ApplyTextLayoutSettings(List<PowerPoint.Shape> shapes, TextLayoutSettings settings)
        {
            if (settings == null) return 0;

            int successCount = 0;

            foreach (var shape in shapes)
            {
                var success = ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return false;

                        var tf = shape.TextFrame;

                        // 余白設定
                        if (settings.MarginLeft.HasValue)
                            tf.MarginLeft = settings.MarginLeft.Value;

                        if (settings.MarginRight.HasValue)
                            tf.MarginRight = settings.MarginRight.Value;

                        if (settings.MarginTop.HasValue)
                            tf.MarginTop = settings.MarginTop.Value;

                        if (settings.MarginBottom.HasValue)
                            tf.MarginBottom = settings.MarginBottom.Value;

                        // 行間設定（TextFrame2経由で設定）
                        if (settings.LineSpacingPt.HasValue)
                        {
                            try
                            {
                                var tf2 = shape.TextFrame2;
                                var paraFmt = tf2.TextRange.ParagraphFormat;
                                paraFmt.SpaceWithin = settings.LineSpacingPt.Value;
                            }
                            catch
                            {
                                // TextFrame2が使えない場合はスキップ
                            }
                        }

                        return true;
                    },
                    $"レイアウト設定適用: {shape.Name}",
                    defaultValue: false,
                    suppressErrors: true);

                if (success) successCount++;
            }

            ComExceptionHandler.LogDebug($"レイアウト設定適用完了: {successCount}/{shapes.Count}個");
            return successCount;
        }

        #endregion

        #region ヘルパーメソッド

        /// <summary>
        /// 大文字小文字を区別しない文字列置換
        /// </summary>
        private string ReplaceIgnoreCase(string source, string search, string replace)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(search))
                return source;

            int index = source.IndexOf(search, StringComparison.OrdinalIgnoreCase);
            if (index < 0) return source;

            var result = new System.Text.StringBuilder();
            int lastIndex = 0;

            while (index >= 0)
            {
                result.Append(source, lastIndex, index - lastIndex);
                result.Append(replace);
                lastIndex = index + search.Length;
                index = source.IndexOf(search, lastIndex, StringComparison.OrdinalIgnoreCase);
            }

            result.Append(source, lastIndex, source.Length - lastIndex);
            return result.ToString();
        }

        #endregion
    }

    /// <summary>
    /// 図形のテキスト情報を格納するクラス
    /// </summary>
    public class ShapeTextInfo
    {
        /// <summary>対象図形</summary>
        public PowerPoint.Shape Shape { get; set; }

        /// <summary>図形名</summary>
        public string ShapeName { get; set; }

        /// <summary>テキスト内容</summary>
        public string Text { get; set; } = "";

        /// <summary>テキストフレームを持つか</summary>
        public bool HasTextFrame { get; set; }

        /// <summary>テキストが存在するか</summary>
        public bool HasText { get; set; }
    }

    /// <summary>
    /// フォント設定クラス（nullのプロパティは変更しない）
    /// </summary>
    public class FontSettings
    {
        /// <summary>フォント名（nullの場合変更しない）</summary>
        public string FontName { get; set; }

        /// <summary>フォントサイズ（nullの場合変更しない）</summary>
        public float? FontSize { get; set; }

        /// <summary>太字（nullの場合変更しない）</summary>
        public bool? IsBold { get; set; }

        /// <summary>斜体（nullの場合変更しない）</summary>
        public bool? IsItalic { get; set; }

        /// <summary>下線（nullの場合変更しない）</summary>
        public bool? IsUnderline { get; set; }

        /// <summary>フォント色RGB（nullの場合変更しない）</summary>
        public int? FontColor { get; set; }

        /// <summary>何か変更設定があるか</summary>
        public bool HasAnySettings =>
            !string.IsNullOrEmpty(FontName) ||
            FontSize.HasValue ||
            IsBold.HasValue ||
            IsItalic.HasValue ||
            IsUnderline.HasValue ||
            FontColor.HasValue;
    }

    /// <summary>
    /// テキストレイアウト設定クラス（nullのプロパティは変更しない）
    /// </summary>
    public class TextLayoutSettings
    {
        /// <summary>行間（pt）（nullの場合変更しない）</summary>
        public float? LineSpacingPt { get; set; }

        /// <summary>左余白（pt）（nullの場合変更しない）</summary>
        public float? MarginLeft { get; set; }

        /// <summary>右余白（pt）（nullの場合変更しない）</summary>
        public float? MarginRight { get; set; }

        /// <summary>上余白（pt）（nullの場合変更しない）</summary>
        public float? MarginTop { get; set; }

        /// <summary>下余白（pt）（nullの場合変更しない）</summary>
        public float? MarginBottom { get; set; }

        /// <summary>何か変更設定があるか</summary>
        public bool HasAnySettings =>
            LineSpacingPt.HasValue ||
            MarginLeft.HasValue ||
            MarginRight.HasValue ||
            MarginTop.HasValue ||
            MarginBottom.HasValue;
    }
}
