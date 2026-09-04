using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// スライド内シェイプの色（塗り・線・フォント）を収集・一括置換するクラス
    /// </summary>
    public class ColorReplacer
    {
        /// <summary>
        /// 指定範囲で使用されている色を収集し、使用回数の多い順に返す
        /// </summary>
        public List<ColorUsageInfo> CollectUsedColors(ColorReplaceScope scope)
        {
            var counts = new Dictionary<int, int>();

            WalkScope(scope, shape =>
            {
                foreach (int rgb in ExtractShapeColors(shape))
                {
                    counts[rgb] = counts.TryGetValue(rgb, out int c) ? c + 1 : 1;
                }
            });

            return counts
                .Select(kv => new ColorUsageInfo { RgbColor = kv.Key, UsageCount = kv.Value })
                .OrderByDescending(c => c.UsageCount)
                .ThenBy(c => c.RgbColor)
                .ToList();
        }

        /// <summary>
        /// 置換マッピングに従って、指定範囲のシェイプの色を一括置換する
        /// </summary>
        /// <param name="scope">対象範囲</param>
        /// <param name="replacementMap">置換前RGB→置換後RGBのマッピング</param>
        /// <returns>変更のあったシェイプ数と、変更したプロパティ（塗り・線・フォントRun）の総数</returns>
        public (int ShapeCount, int PropertyCount) ApplyReplacements(ColorReplaceScope scope, Dictionary<int, int> replacementMap)
        {
            if (replacementMap == null || replacementMap.Count == 0)
                return (0, 0);

            int shapeCount = 0;
            int propertyCount = 0;

            WalkScope(scope, shape =>
            {
                int changed = ApplyToShape(shape, replacementMap);
                if (changed > 0)
                {
                    shapeCount++;
                    propertyCount += changed;
                }
            });

            return (shapeCount, propertyCount);
        }

        #region スコープ走査（グループ内を再帰的に処理）

        private void WalkScope(ColorReplaceScope scope, Action<PowerPoint.Shape> action)
        {
            var app = Globals.ThisAddIn.Application;

            if (scope == ColorReplaceScope.CurrentSlide)
            {
                var slide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
                if (slide == null) return;
                WalkShapes(slide.Shapes, action);
            }
            else
            {
                var presentation = app?.ActivePresentation;
                if (presentation == null) return;
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    WalkShapes(slide.Shapes, action);
                }
            }
        }

        private void WalkShapes(IEnumerable shapes, Action<PowerPoint.Shape> action)
        {
            if (shapes == null) return;

            foreach (PowerPoint.Shape shape in shapes)
            {
                ComExceptionHandler.ExecuteComOperation(() =>
                {
                    if (shape.Type == Office.MsoShapeType.msoGroup)
                    {
                        // グループ自体の色は対象とせず、子シェイプのみ再帰的に処理する
                        WalkShapes(shape.GroupItems, action);
                    }
                    else
                    {
                        action(shape);
                    }
                }, $"色置換走査: {SafeShapeName(shape)}", suppressErrors: true);
            }
        }

        #endregion

        #region 色抽出・適用

        /// <summary>単色塗り・線・フォント(Run単位)の色をすべて列挙する</summary>
        private static IEnumerable<int> ExtractShapeColors(PowerPoint.Shape shape)
        {
            var colors = new List<int>();

            try
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue &&
                    shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                {
                    colors.Add(shape.Fill.ForeColor.RGB);
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"塗り色取得失敗 [{SafeShapeName(shape)}]: {ex.Message}");
            }

            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    colors.Add(shape.Line.ForeColor.RGB);
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"線色取得失敗 [{SafeShapeName(shape)}]: {ex.Message}");
            }

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    colors.AddRange(ExtractTextFrameColors(shape.TextFrame));
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"フォント色取得失敗 [{SafeShapeName(shape)}]: {ex.Message}");
            }

            try
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    ForEachTableCell(shape.Table, cellShape =>
                    {
                        if (cellShape.Fill.Visible == Office.MsoTriState.msoTrue &&
                            cellShape.Fill.Type == Office.MsoFillType.msoFillSolid)
                        {
                            colors.Add(cellShape.Fill.ForeColor.RGB);
                        }
                        if (cellShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            colors.AddRange(ExtractTextFrameColors(cellShape.TextFrame));
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                ComExceptionHandler.LogWarning($"表セル色取得失敗 [{SafeShapeName(shape)}]: {ex.Message}");
            }

            return colors;
        }

        /// <summary>テキストフレーム内のRun単位のフォント色をすべて列挙する</summary>
        private static IEnumerable<int> ExtractTextFrameColors(PowerPoint.TextFrame textFrame)
        {
            var colors = new List<int>();
            var textRange = textFrame.TextRange;
            int runCount = textRange.Runs().Count;
            for (int i = 1; i <= runCount; i++)
            {
                colors.Add(textRange.Runs(i, 1).Font.Color.RGB);
            }
            return colors;
        }

        /// <summary>表の全セルのシェイプに対してactionを実行する</summary>
        private static void ForEachTableCell(PowerPoint.Table table, Action<PowerPoint.Shape> action)
        {
            for (int r = 1; r <= table.Rows.Count; r++)
            {
                for (int c = 1; c <= table.Columns.Count; c++)
                {
                    action(table.Cell(r, c).Shape);
                }
            }
        }

        /// <summary>replacementMapに一致する色を置換し、変更したプロパティ数を返す</summary>
        private static int ApplyToShape(PowerPoint.Shape shape, Dictionary<int, int> replacementMap)
        {
            int changed = 0;

            changed += ComExceptionHandler.ExecuteComOperation(() =>
            {
                int c = 0;
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue &&
                    shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                {
                    int rgb = shape.Fill.ForeColor.RGB;
                    if (replacementMap.TryGetValue(rgb, out int newRgb) && newRgb != rgb)
                    {
                        shape.Fill.ForeColor.RGB = newRgb;
                        c = 1;
                    }
                }
                return c;
            }, $"塗り色置換: {SafeShapeName(shape)}", defaultValue: 0, suppressErrors: true);

            changed += ComExceptionHandler.ExecuteComOperation(() =>
            {
                int c = 0;
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    int rgb = shape.Line.ForeColor.RGB;
                    if (replacementMap.TryGetValue(rgb, out int newRgb) && newRgb != rgb)
                    {
                        shape.Line.ForeColor.RGB = newRgb;
                        c = 1;
                    }
                }
                return c;
            }, $"線色置換: {SafeShapeName(shape)}", defaultValue: 0, suppressErrors: true);

            changed += ComExceptionHandler.ExecuteComOperation(() =>
            {
                int c = 0;
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    c += ApplyTextFrameReplacements(shape.TextFrame, replacementMap);
                }
                return c;
            }, $"フォント色置換: {SafeShapeName(shape)}", defaultValue: 0, suppressErrors: true);

            changed += ComExceptionHandler.ExecuteComOperation(() =>
            {
                int c = 0;
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    ForEachTableCell(shape.Table, cellShape =>
                    {
                        if (cellShape.Fill.Visible == Office.MsoTriState.msoTrue &&
                            cellShape.Fill.Type == Office.MsoFillType.msoFillSolid)
                        {
                            int rgb = cellShape.Fill.ForeColor.RGB;
                            if (replacementMap.TryGetValue(rgb, out int newRgb) && newRgb != rgb)
                            {
                                cellShape.Fill.ForeColor.RGB = newRgb;
                                c++;
                            }
                        }
                        if (cellShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            c += ApplyTextFrameReplacements(cellShape.TextFrame, replacementMap);
                        }
                    });
                }
                return c;
            }, $"表セル色置換: {SafeShapeName(shape)}", defaultValue: 0, suppressErrors: true);

            return changed;
        }

        /// <summary>replacementMapに一致するテキストフレーム内のRunフォント色を置換し、変更数を返す</summary>
        private static int ApplyTextFrameReplacements(PowerPoint.TextFrame textFrame, Dictionary<int, int> replacementMap)
        {
            int c = 0;
            var textRange = textFrame.TextRange;
            int runCount = textRange.Runs().Count;
            for (int i = 1; i <= runCount; i++)
            {
                var run = textRange.Runs(i, 1);
                int rgb = run.Font.Color.RGB;
                if (replacementMap.TryGetValue(rgb, out int newRgb) && newRgb != rgb)
                {
                    run.Font.Color.RGB = newRgb;
                    c++;
                }
            }
            return c;
        }

        private static string SafeShapeName(PowerPoint.Shape shape)
        {
            try { return shape.Name; }
            catch { return "(不明)"; }
        }

        #endregion
    }
}
