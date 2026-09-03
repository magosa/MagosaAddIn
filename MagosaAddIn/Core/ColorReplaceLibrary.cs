using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 色置換リストのライブラリ管理クラス
    /// 名前付きの置換色リストをJSONファイルに永続化する
    /// </summary>
    public class ColorReplaceLibrary
    {
        #region 定数

        private const string AppName = "MagosaAddIn";
        private const string FileName = "ColorReplaceLibrary.json";
        private const int MaxListCount = 50;

        #endregion

        #region フィールド

        private List<ColorReplaceListEntry> _lists;
        private readonly string _filePath;

        #endregion

        #region コンストラクタ

        public ColorReplaceLibrary()
        {
            _filePath = GetSaveFilePath();
            _lists = new List<ColorReplaceListEntry>();
            LoadFromFile();
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// 置換色リストを名前を付けて保存（同名が存在する場合は上書き）
        /// </summary>
        public void SaveList(string name, List<ColorReplacementEntry> entries)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("リスト名を入力してください");
            if (entries == null) throw new ArgumentNullException(nameof(entries));

            var existing = _lists.FirstOrDefault(l => l.Name == name);
            if (existing != null)
            {
                _lists.Remove(existing);
            }
            else if (_lists.Count >= MaxListCount)
            {
                throw new InvalidOperationException($"保存できるリストの最大数({MaxListCount})に達しています。不要なリストを削除してください。");
            }

            _lists.Add(new ColorReplaceListEntry
            {
                Name = name,
                CreatedAt = DateTime.Now.ToString("yyyy/MM/dd HH:mm"),
                Entries = entries.Select(e => new ColorReplacementEntry
                {
                    OriginalColor = e.OriginalColor,
                    ReplacementColor = e.ReplacementColor
                }).ToList()
            });

            SaveToFile();
            ComExceptionHandler.LogDebug($"色置換リスト保存: '{name}' ({entries.Count}件)");
        }

        /// <summary>
        /// 指定名のリストを取得（見つからない場合はnull）
        /// </summary>
        public List<ColorReplacementEntry> LoadList(string name)
        {
            return _lists.FirstOrDefault(l => l.Name == name)?.Entries;
        }

        /// <summary>
        /// 保存済みリストを全件取得
        /// </summary>
        public List<ColorReplaceListEntry> GetAllLists() => _lists.ToList();

        /// <summary>
        /// 指定名のリストを削除
        /// </summary>
        public bool DeleteList(string name)
        {
            int removed = _lists.RemoveAll(l => l.Name == name);
            if (removed > 0) SaveToFile();
            return removed > 0;
        }

        /// <summary>
        /// 名前の重複チェック
        /// </summary>
        public bool ExistsName(string name) => _lists.Any(l => l.Name == name);

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
                var data = new ColorReplaceLibraryData { Lists = _lists };
                string json = SerializeToJson(data);
                File.WriteAllText(_filePath, json, Encoding.UTF8);
                ComExceptionHandler.LogDebug($"色置換リストライブラリ保存: {_lists.Count}件 → {_filePath}");
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("色置換リストライブラリ保存失敗", ex);
            }
        }

        private void LoadFromFile()
        {
            try
            {
                if (!File.Exists(_filePath))
                {
                    _lists = new List<ColorReplaceListEntry>();
                    return;
                }

                string json = File.ReadAllText(_filePath, Encoding.UTF8);
                var data = DeserializeFromJson(json);
                _lists = data?.Lists ?? new List<ColorReplaceListEntry>();
                ComExceptionHandler.LogDebug($"色置換リストライブラリ読み込み: {_lists.Count}件");
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogError("色置換リストライブラリ読み込み失敗", ex);
                _lists = new List<ColorReplaceListEntry>();
            }
        }

        private static string SerializeToJson(ColorReplaceLibraryData data)
        {
            var serializer = new DataContractJsonSerializer(typeof(ColorReplaceLibraryData));
            using (var ms = new MemoryStream())
            {
                serializer.WriteObject(ms, data);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private static ColorReplaceLibraryData DeserializeFromJson(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(ColorReplaceLibraryData));
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (ColorReplaceLibraryData)serializer.ReadObject(ms);
            }
        }

        #endregion
    }

    /// <summary>
    /// 色置換リストライブラリのルート要素（JSONシリアライズ用）
    /// </summary>
    [DataContract]
    public class ColorReplaceLibraryData
    {
        [DataMember] public List<ColorReplaceListEntry> Lists { get; set; } = new List<ColorReplaceListEntry>();
    }
}
