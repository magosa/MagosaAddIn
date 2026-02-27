using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形スタイルライブラリ管理クラス
    /// スタイルをJSONファイルに永続化する
    /// </summary>
    public class ShapeStyleLibrary
    {
        #region 定数

        private const string AppName = "MagosaAddIn";
        private const string FileName = "StyleLibrary.json";
        private const int MaxStyleCount = 100;

        #endregion

        #region フィールド

        private List<StyleEntry> _styles;
        private readonly string _filePath;

        #endregion

        #region コンストラクタ

        public ShapeStyleLibrary()
        {
            _filePath = GetSaveFilePath();
            _styles = new List<StyleEntry>();
            LoadFromFile();
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// 図形からスタイルを保存
        /// </summary>
        public StyleEntry SaveStyleFromShape(PowerPoint.Shape shape, string name)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("スタイル名を入力してください");
            if (_styles.Count >= MaxStyleCount)
                throw new InvalidOperationException($"スタイルの最大数({MaxStyleCount})に達しています。不要なスタイルを削除してください。");

            var entry = ComExceptionHandler.ExecuteComOperation(
                () => ExtractStyle(shape, name),
                "スタイル保存");

            _styles.Add(entry);
            SaveToFile();
            ComExceptionHandler.LogDebug($"スタイル保存: '{name}'");
            return entry;
        }

        /// <summary>
        /// スタイルを図形に適用
        /// </summary>
        public int ApplyStyleToShapes(List<PowerPoint.Shape> shapes, string styleName)
        {
            var entry = _styles.FirstOrDefault(s => s.Name == styleName);
            if (entry == null) throw new InvalidOperationException($"スタイル '{styleName}' が見つかりません");

            int count = 0;
            foreach (var shape in shapes)
            {
                bool ok = ComExceptionHandler.ExecuteComOperation(
                    () => { ApplyStyleToShape(shape, entry); return true; },
                    $"スタイル適用: {shape.Name}",
                    defaultValue: false,
                    suppressErrors: true);
                if (ok) count++;
            }
            ComExceptionHandler.LogDebug($"スタイル適用完了: '{styleName}' → {count}個");
            return count;
        }

        /// <summary>
        /// 指定スタイルを削除
        /// </summary>
        public bool DeleteStyle(string name)
        {
            int removed = _styles.RemoveAll(s => s.Name == name);
            if (removed > 0) SaveToFile();
            return removed > 0;
        }

        /// <summary>
        /// お気に入り切り替え
        /// </summary>
        public void ToggleFavorite(string name)
        {
            var entry = _styles.FirstOrDefault(s => s.Name == name);
            if (entry != null)
            {
                entry.IsFavorite = !entry.IsFavorite;
                SaveToFile();
            }
        }

        /// <summary>
        /// 全スタイル取得
        /// </summary>
        public List<StyleEntry> GetAllStyles() => _styles.ToList();

        /// <summary>
        /// お気に入りスタイル取得
        /// </summary>
        public List<StyleEntry> GetFavoriteStyles() => _styles.Where(s => s.IsFavorite).ToList();

        /// <summary>
        /// スタイル数取得
        /// </summary>
        public int Count => _styles.Count;

        /// <summary>
        /// 名前の重複チェック
        /// </summary>
        public bool ExistsName(string name) => _styles.Any(s => s.Name == name);

        /// <summary>
        /// JSONエクスポート文字列取得
        /// </summary>
        public string ExportToJson()
        {
            var data = new StyleLibraryData { Styles = _styles };
            return SerializeToJson(data);
        }

        /// <summary>
        /// JSONからインポート（既存スタイルにマージ）
        /// </summary>
        public int ImportFromJson(string json, bool overwrite = false)
        {
            var data = DeserializeFromJson(json);
            if (data?.Styles == null) return 0;

            int count = 0;
            foreach (var entry in data.Styles)
            {
                var existing = _styles.FirstOrDefault(s => s.Name == entry.Name);
                if (existing != null)
                {
                    if (overwrite)
                    {
                        _styles.Remove(existing);
                        _styles.Add(entry);
                        count++;
                    }
                }
                else if (_styles.Count < MaxStyleCount)
                {
                    _styles.Add(entry);
                    count++;
                }
            }

            if (count > 0) SaveToFile();
            return count;
        }

        #endregion

        #region スタイル抽出・適用

        private StyleEntry ExtractStyle(PowerPoint.Shape shape, string name)
        {
            var entry = new StyleEntry
            {
                Name = name,
                CreatedAt = DateTime.Now.ToString("yyyy/MM/dd HH:mm"),
                IsFavorite = false
            };

            // 塗りつぶし
            try
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    entry.HasFill = true;
                    entry.FillType = (int)shape.Fill.Type;

                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                    {
                        entry.FillColor = shape.Fill.ForeColor.RGB;
                        entry.FillTransparency = shape.Fill.Transparency;
                    }
                    else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                    {
                        entry.HasGradient = true;
                        entry.GradientColor1 = shape.Fill.ForeColor.RGB;
                        entry.GradientColor2 = shape.Fill.BackColor.RGB;
                        entry.GradientAngle = shape.Fill.GradientAngle;
                        entry.FillTransparency = shape.Fill.Transparency;
                    }
                }
                else
                {
                    entry.HasFill = false;
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"塗りつぶし情報取得失敗: {ex.Message}");
            }

            // 枠線
            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    entry.HasLine = true;
                    entry.LineColor = shape.Line.ForeColor.RGB;
                    entry.LineWeight = shape.Line.Weight;
                    entry.LineDashStyle = (int)shape.Line.DashStyle;
                }
                else
                {
                    entry.HasLine = false;
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"枠線情報取得失敗: {ex.Message}");
            }

            // 影
            try
            {
                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                {
                    entry.HasShadow = true;
                    entry.ShadowColor = shape.Shadow.ForeColor.RGB;
                    entry.ShadowTransparency = shape.Shadow.Transparency;
                    entry.ShadowOffsetX = shape.Shadow.OffsetX;
                    entry.ShadowOffsetY = shape.Shadow.OffsetY;
                    entry.ShadowSize = shape.Shadow.Size;
                    entry.ShadowBlur = shape.Shadow.Blur;
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"影情報取得失敗: {ex.Message}");
            }

            // フォント
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var tf = shape.TextFrame.TextRange;
                    entry.FontName = tf.Font.Name;
                    entry.FontSize = tf.Font.Size;
                    entry.FontBold = tf.Font.Bold == Office.MsoTriState.msoTrue;
                    entry.FontItalic = tf.Font.Italic == Office.MsoTriState.msoTrue;
                    entry.FontColor = tf.Font.Color.RGB;
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"フォント情報取得失敗: {ex.Message}");
            }

            return entry;
        }

        private void ApplyStyleToShape(PowerPoint.Shape shape, StyleEntry entry)
        {
            // 塗りつぶし
            if (entry.HasFill)
            {
                try
                {
                    if (entry.HasGradient)
                    {
                        // グラデーションは簡易的に前面色のみ適用
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.ForeColor.RGB = entry.GradientColor1;
                        shape.Fill.Transparency = entry.FillTransparency;
                    }
                    else
                    {
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = entry.FillColor;
                        shape.Fill.Transparency = entry.FillTransparency;
                    }
                }
                catch (Exception ex) { ComExceptionHandler.LogWarning($"塗りつぶし適用失敗: {ex.Message}"); }
            }
            else
            {
                try { shape.Fill.Visible = Office.MsoTriState.msoFalse; }
                catch { }
            }

            // 枠線
            if (entry.HasLine)
            {
                try
                {
                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = entry.LineColor;
                    shape.Line.Weight = entry.LineWeight;
                    shape.Line.DashStyle = (Office.MsoLineDashStyle)entry.LineDashStyle;
                }
                catch (Exception ex) { ComExceptionHandler.LogWarning($"枠線適用失敗: {ex.Message}"); }
            }
            else
            {
                try { shape.Line.Visible = Office.MsoTriState.msoFalse; }
                catch { }
            }

            // 影
            if (entry.HasShadow)
            {
                try
                {
                    shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                    shape.Shadow.ForeColor.RGB = entry.ShadowColor;
                    shape.Shadow.Transparency = entry.ShadowTransparency;
                    shape.Shadow.OffsetX = entry.ShadowOffsetX;
                    shape.Shadow.OffsetY = entry.ShadowOffsetY;
                    if (entry.ShadowSize > 0) shape.Shadow.Size = entry.ShadowSize;
                    if (entry.ShadowBlur > 0) shape.Shadow.Blur = entry.ShadowBlur;
                }
                catch (Exception ex) { ComExceptionHandler.LogWarning($"影適用失敗: {ex.Message}"); }
            }
            else
            {
                try { shape.Shadow.Visible = Office.MsoTriState.msoFalse; }
                catch { }
            }

            // フォント
            if (!string.IsNullOrEmpty(entry.FontName))
            {
                try
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        var tf = shape.TextFrame.TextRange;
                        if (!string.IsNullOrEmpty(entry.FontName)) tf.Font.Name = entry.FontName;
                        if (entry.FontSize > 0) tf.Font.Size = entry.FontSize;
                        tf.Font.Bold = entry.FontBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                        tf.Font.Italic = entry.FontItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                        if (entry.FontColor > 0) tf.Font.Color.RGB = entry.FontColor;
                    }
                }
                catch (Exception ex) { ComExceptionHandler.LogWarning($"フォント適用失敗: {ex.Message}"); }
            }
        }

        #endregion

        #region 永続化（JSON）

        private string GetSaveFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dir = Path.Combine(appData, AppName);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, FileName);
        }

        private void SaveToFile()
        {
            try
            {
                var data = new StyleLibraryData { Styles = _styles };
                string json = SerializeToJson(data);
                File.WriteAllText(_filePath, json, Encoding.UTF8);
                ComExceptionHandler.LogDebug($"スタイルライブラリ保存: {_styles.Count}件 → {_filePath}");
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("スタイルライブラリ保存失敗", ex);
            }
        }

        private void LoadFromFile()
        {
            try
            {
                if (!File.Exists(_filePath))
                {
                    _styles = new List<StyleEntry>();
                    return;
                }

                string json = File.ReadAllText(_filePath, Encoding.UTF8);
                var data = DeserializeFromJson(json);
                _styles = data?.Styles ?? new List<StyleEntry>();
                ComExceptionHandler.LogDebug($"スタイルライブラリ読み込み: {_styles.Count}件");
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("スタイルライブラリ読み込み失敗", ex);
                _styles = new List<StyleEntry>();
            }
        }

        private static string SerializeToJson(StyleLibraryData data)
        {
            var serializer = new DataContractJsonSerializer(typeof(StyleLibraryData));
            using (var ms = new MemoryStream())
            {
                serializer.WriteObject(ms, data);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private static StyleLibraryData DeserializeFromJson(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(StyleLibraryData));
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (StyleLibraryData)serializer.ReadObject(ms);
            }
        }

        #endregion
    }

    /// <summary>
    /// スタイルエントリ（1件のスタイル情報）
    /// </summary>
    [DataContract]
    public class StyleEntry
    {
        [DataMember] public string Name { get; set; }
        [DataMember] public string CreatedAt { get; set; }
        [DataMember] public bool IsFavorite { get; set; }

        // 塗りつぶし
        [DataMember] public bool HasFill { get; set; }
        [DataMember] public int FillType { get; set; }
        [DataMember] public int FillColor { get; set; }
        [DataMember] public float FillTransparency { get; set; }
        [DataMember] public bool HasGradient { get; set; }
        [DataMember] public int GradientColor1 { get; set; }
        [DataMember] public int GradientColor2 { get; set; }
        [DataMember] public float GradientAngle { get; set; }

        // 枠線
        [DataMember] public bool HasLine { get; set; }
        [DataMember] public int LineColor { get; set; }
        [DataMember] public float LineWeight { get; set; }
        [DataMember] public int LineDashStyle { get; set; }

        // 影
        [DataMember] public bool HasShadow { get; set; }
        [DataMember] public int ShadowColor { get; set; }
        [DataMember] public float ShadowTransparency { get; set; }
        [DataMember] public float ShadowOffsetX { get; set; }
        [DataMember] public float ShadowOffsetY { get; set; }
        [DataMember] public float ShadowSize { get; set; }
        [DataMember] public float ShadowBlur { get; set; }

        // フォント
        [DataMember] public string FontName { get; set; }
        [DataMember] public float FontSize { get; set; }
        [DataMember] public bool FontBold { get; set; }
        [DataMember] public bool FontItalic { get; set; }
        [DataMember] public int FontColor { get; set; }

        /// <summary>
        /// 塗りつぶし色（表示用 System.Drawing.Color に変換）
        /// </summary>
        public System.Drawing.Color GetFillDrawingColor()
        {
            if (!HasFill) return System.Drawing.Color.Transparent;
            int rgb = HasGradient ? GradientColor1 : FillColor;
            // PowerPoint RGB: B<<16 | G<<8 | R
            return System.Drawing.Color.FromArgb(
                rgb & 0xFF,
                (rgb >> 8) & 0xFF,
                (rgb >> 16) & 0xFF);
        }

        /// <summary>
        /// 枠線色（表示用）
        /// </summary>
        public System.Drawing.Color GetLineDrawingColor()
        {
            if (!HasLine) return System.Drawing.Color.Transparent;
            return System.Drawing.Color.FromArgb(
                LineColor & 0xFF,
                (LineColor >> 8) & 0xFF,
                (LineColor >> 16) & 0xFF);
        }

        /// <summary>
        /// スタイルのサマリーテキスト
        /// </summary>
        public string GetSummary()
        {
            var parts = new List<string>();
            if (HasFill)
                parts.Add(HasGradient ? "グラデーション" : $"塗り #{FillColor:X6}");
            else
                parts.Add("塗りなし");

            if (HasLine)
                parts.Add($"枠 {LineWeight:F1}pt");
            else
                parts.Add("枠なし");

            if (HasShadow) parts.Add("影あり");
            if (!string.IsNullOrEmpty(FontName)) parts.Add($"{FontName} {FontSize:F0}pt");

            return string.Join(" ／ ", parts);
        }
    }

    /// <summary>
    /// ライブラリデータコンテナ（JSON直列化用）
    /// </summary>
    [DataContract]
    public class StyleLibraryData
    {
        [DataMember]
        public List<StyleEntry> Styles { get; set; } = new List<StyleEntry>();
    }
}
