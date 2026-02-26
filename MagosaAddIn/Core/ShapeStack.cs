using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形スタック管理クラス（複数スタック対応）
    /// </summary>
    public class ShapeStack
    {
        private Dictionary<int, List<PowerPoint.Shape>> stacks;
        private Dictionary<int, DateTime> stackCreatedAt;
        private int nextStackId;

        public ShapeStack()
        {
            stacks = new Dictionary<int, List<PowerPoint.Shape>>();
            stackCreatedAt = new Dictionary<int, DateTime>();
            nextStackId = 1;
        }

        /// <summary>
        /// 新しいスタックを追加
        /// </summary>
        /// <param name="shapes">スタックする図形リスト</param>
        /// <returns>作成されたスタックID</returns>
        public int PushStack(List<PowerPoint.Shape> shapes)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (shapes == null || shapes.Count == 0)
                    {
                        throw new ArgumentException("スタックする図形が指定されていません。");
                    }

                    // 最大スタック数チェック
                    if (stacks.Count >= Constants.MAX_STACK_COUNT)
                    {
                        throw new InvalidOperationException($"スタックの最大数({Constants.MAX_STACK_COUNT})に達しています。不要なスタックを削除してください。");
                    }

                    int stackId = nextStackId++;
                    stacks[stackId] = new List<PowerPoint.Shape>(shapes);
                    stackCreatedAt[stackId] = DateTime.Now;

                    ComExceptionHandler.LogDebug($"スタック追加: Stack {stackId}, 図形数: {shapes.Count}");
                    return stackId;
                },
                "スタック追加");
        }

        /// <summary>
        /// 指定スタックを復元（図形を選択状態にする）
        /// </summary>
        /// <param name="stackId">スタックID</param>
        /// <returns>復元成功したかどうか</returns>
        public bool RestoreStack(int stackId)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (!stacks.ContainsKey(stackId))
                    {
                        ComExceptionHandler.LogWarning($"Stack {stackId}が見つかりません。");
                        return false;
                    }

                    var shapes = stacks[stackId];
                    if (shapes == null || shapes.Count == 0)
                    {
                        ComExceptionHandler.LogWarning($"Stack {stackId}に図形が含まれていません。");
                        return false;
                    }

                    // 有効な図形のみをフィルタリング
                    var validShapes = new List<PowerPoint.Shape>();
                    foreach (var shape in shapes)
                    {
                        try
                        {
                            // 図形が有効かチェック（削除されていないか）
                            var name = shape.Name; // アクセスできればOK
                            validShapes.Add(shape);
                        }
                        catch
                        {
                            ComExceptionHandler.LogDebug($"Stack {stackId}: 無効な図形をスキップ");
                        }
                    }

                    if (validShapes.Count == 0)
                    {
                        ComExceptionHandler.LogWarning($"Stack {stackId}の図形はすべて削除されています。");
                        return false;
                    }

                    // PowerPointで図形を選択
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow == null)
                    {
                        throw new InvalidOperationException("アクティブウィンドウが見つかりません。");
                    }

                    // 図形の名前を配列として取得
                    var shapeNames = new object[validShapes.Count];
                    for (int i = 0; i < validShapes.Count; i++)
                    {
                        shapeNames[i] = validShapes[i].Name;
                    }

                    // スライドを取得
                    var slide = validShapes[0].Parent as PowerPoint.Slide;
                    if (slide == null)
                    {
                        throw new InvalidOperationException("図形が所属するスライドを取得できません。");
                    }

                    // ShapeRangeを使って一度に複数の図形を選択
                    var shapeRange = slide.Shapes.Range(shapeNames);
                    shapeRange.Select(Microsoft.Office.Core.MsoTriState.msoFalse);

                    ComExceptionHandler.LogDebug($"スタック復元: Stack {stackId}, {validShapes.Count}個の図形を選択");
                    return true;
                },
                $"スタック復元 (Stack {stackId})",
                defaultValue: false,
                suppressErrors: true);
        }

        /// <summary>
        /// 指定スタックを削除
        /// </summary>
        /// <param name="stackId">スタックID</param>
        public void RemoveStack(int stackId)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (stacks.ContainsKey(stackId))
                    {
                        stacks.Remove(stackId);
                        stackCreatedAt.Remove(stackId);
                        ComExceptionHandler.LogDebug($"スタック削除: Stack {stackId}");
                    }
                },
                "スタック削除",
                suppressErrors: true);
        }

        /// <summary>
        /// すべてのスタックをクリア
        /// </summary>
        public void ClearAllStacks()
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    stacks.Clear();
                    stackCreatedAt.Clear();
                    nextStackId = 1;
                    ComExceptionHandler.LogDebug("すべてのスタックをクリアしました");
                },
                "スタック全クリア",
                suppressErrors: true);
        }

        /// <summary>
        /// スタック情報一覧を取得
        /// </summary>
        /// <returns>スタック情報リスト</returns>
        public List<StackInfo> GetStackInfoList()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    var infoList = new List<StackInfo>();

                    foreach (var kvp in stacks.OrderBy(x => x.Key))
                    {
                        int stackId = kvp.Key;
                        var shapes = kvp.Value;

                        // 有効な図形数をカウント
                        int validCount = 0;
                        foreach (var shape in shapes)
                        {
                            try
                            {
                                var name = shape.Name;
                                validCount++;
                            }
                            catch
                            {
                                // 無効な図形
                            }
                        }

                        var info = new StackInfo
                        {
                            StackId = stackId,
                            ShapeCount = validCount,
                            CreatedAt = stackCreatedAt.ContainsKey(stackId) ? stackCreatedAt[stackId] : DateTime.MinValue,
                            IsValid = validCount > 0
                        };

                        infoList.Add(info);
                    }

                    return infoList;
                },
                "スタック情報取得",
                defaultValue: new List<StackInfo>(),
                suppressErrors: true);
        }

        /// <summary>
        /// 総スタック数を取得
        /// </summary>
        /// <returns>スタック数</returns>
        public int GetTotalStackCount()
        {
            return stacks.Count;
        }

        /// <summary>
        /// 総図形数を取得（有効な図形のみ）
        /// </summary>
        /// <returns>図形数</returns>
        public int GetTotalShapeCount()
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    int totalCount = 0;
                    foreach (var shapes in stacks.Values)
                    {
                        foreach (var shape in shapes)
                        {
                            try
                            {
                                var name = shape.Name;
                                totalCount++;
                            }
                            catch
                            {
                                // 無効な図形はカウントしない
                            }
                        }
                    }
                    return totalCount;
                },
                "総図形数取得",
                defaultValue: 0,
                suppressErrors: true);
        }
    }

    /// <summary>
    /// スタック情報
    /// </summary>
    public class StackInfo
    {
        /// <summary>
        /// スタックID
        /// </summary>
        public int StackId { get; set; }

        /// <summary>
        /// 図形数（有効な図形のみ）
        /// </summary>
        public int ShapeCount { get; set; }

        /// <summary>
        /// 作成日時
        /// </summary>
        public DateTime CreatedAt { get; set; }

        /// <summary>
        /// スタックが有効かどうか（少なくとも1つの有効な図形を含む）
        /// </summary>
        public bool IsValid { get; set; }
    }
}
