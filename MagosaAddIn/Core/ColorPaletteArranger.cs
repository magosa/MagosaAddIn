using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using ColorConv = MagosaAddIn.Core.ColorConverter;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// カラーパレット配置クラス
    /// スライドの枠外にカラーグリッドを配置
    /// </summary>
    public class ColorPaletteArranger
    {
        /// <summary>
        /// カラーグリッドをスライド枠外に配置
        /// </summary>
        /// <param name="colorMatrix">色×明度段階の2次元リスト</param>
        /// <param name="options">配置オプション</param>
        public void ArrangeColorGrid(List<List<int>> colorMatrix, PaletteArrangementOptions options = null)
        {
            if (colorMatrix == null || colorMatrix.Count == 0)
                throw new ArgumentException("カラーマトリックスが空です。");

            var app = Globals.ThisAddIn.Application;
            var slide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
            if (slide == null)
                throw new InvalidOperationException("アクティブなスライドがありません。");

            options = options ?? new PaletteArrangementOptions();

            // スライドサイズ取得
            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

            // 配置位置計算（スライドの枠外に適切なマージンを設けて配置）
            float startX, startY;
            if (options.Position == PalettePosition.Right)
            {
                startX = slideWidth + options.Margin;
                startY = options.Margin;
            }
            else // Bottom
            {
                startX = options.Margin;
                startY = slideHeight + options.Margin;
            }

            // カラーグリッドを作成
            CreateColorGrid(slide, colorMatrix, startX, startY, options.CellSize);
        }

        /// <summary>
        /// カラーグリッドをスライド上に配置（選択図形への適用と同時に実行）
        /// </summary>
        /// <param name="colors">カラーリスト</param>
        /// <param name="options">配置オプション</param>
        public void ArrangeColorRow(List<int> colors, PaletteArrangementOptions options = null)
        {
            if (colors == null || colors.Count == 0)
                throw new ArgumentException("カラーリストが空です。");

            var app = Globals.ThisAddIn.Application;
            var slide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
            if (slide == null)
                throw new InvalidOperationException("アクティブなスライドがありません。");

            options = options ?? new PaletteArrangementOptions();

            // スライドサイズ取得
            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

            // 配置位置計算
            float startX, startY;
            if (options.Position == PalettePosition.Right)
            {
                startX = slideWidth + options.Margin;
                startY = options.Margin;
            }
            else // Bottom
            {
                startX = options.Margin;
                startY = slideHeight + options.Margin;
            }

            // カラー行を作成
            CreateColorRow(slide, colors, startX, startY, options.CellSize);
        }

        /// <summary>
        /// カラーグリッドを作成（色×明度段階）
        /// </summary>
        private void CreateColorGrid(PowerPoint.Slide slide, List<List<int>> colorMatrix, 
            float startX, float startY, float cellSize)
        {
            int colorCount = colorMatrix.Count;
            int lightnessSteps = colorMatrix[0].Count;
            
            // セル間のマージン（1pt）
            const float cellMargin = 1f;
            float cellWithMargin = cellSize + cellMargin;

            for (int col = 0; col < colorCount; col++)
            {
                for (int row = 0; row < lightnessSteps; row++)
                {
                    float x = startX + (col * cellWithMargin);
                    float y = startY + (row * cellWithMargin);
                    int color = colorMatrix[col][row];

                    CreateColorCell(slide, color, x, y, cellSize);
                }
            }
        }

        /// <summary>
        /// カラー行を作成（単一行）
        /// </summary>
        private void CreateColorRow(PowerPoint.Slide slide, List<int> colors, 
            float startX, float startY, float cellSize)
        {
            // セル間のマージン（1pt）
            const float cellMargin = 1f;
            float cellWithMargin = cellSize + cellMargin;
            
            for (int i = 0; i < colors.Count; i++)
            {
                float x = startX + (i * cellWithMargin);
                CreateColorCell(slide, colors[i], x, startY, cellSize);
            }
        }

        /// <summary>
        /// カラーセルを作成（正方形図形 + カラーコードテキスト）
        /// </summary>
        private PowerPoint.Shape CreateColorCell(PowerPoint.Slide slide, int color, 
            float x, float y, float cellSize)
        {
            return ComExceptionHandler.ExecuteComOperation(() =>
            {
                // 正方形図形を作成
                var shape = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    x, y, cellSize, cellSize);

                // 塗りつぶし色を設定
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = color;

                // 枠線を非表示に設定
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                // テキストフレーム設定（カラーコードを表示）
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textFrame = shape.TextFrame;
                    textFrame.TextRange.Text = ColorConv.RgbToHex(color);
                    
                    // フォントサイズを固定（4pt：セル内に収まるサイズ）
                    const float fontSize = 4f;
                    textFrame.TextRange.Font.Size = fontSize;
                    textFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    
                    // テキスト色を背景色に応じて自動調整
                    var (h, s, l) = ColorConv.RgbToHsl(color);
                    const float lightnessThreshold = 0.5f;
                    int textColor = l > lightnessThreshold ? Constants.DEFAULT_LINE_COLOR : Constants.DEFAULT_FILL_COLOR;
                    textFrame.TextRange.Font.Color.RGB = textColor;
                    
                    textFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                    textFrame.TextRange.ParagraphFormat.Alignment = 
                        PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    textFrame.WordWrap = Office.MsoTriState.msoFalse;
                    textFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                }

                return shape;
            }, "カラーセル作成", defaultValue: null);
        }

        /// <summary>
        /// 選択図形に色を適用
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="colors">適用する色のリスト</param>
        /// <param name="applyMode">適用モード（順番/ランダム）</param>
        public void ApplyColorsToShapes(List<PowerPoint.Shape> shapes, List<int> colors, 
            ColorApplyMode applyMode = ColorApplyMode.Sequential)
        {
            if (shapes == null || shapes.Count == 0)
                throw new ArgumentException("図形リストが空です。");
            if (colors == null || colors.Count == 0)
                throw new ArgumentException("カラーリストが空です。");

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                var random = new Random();

                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    int color;

                    switch (applyMode)
                    {
                        case ColorApplyMode.Sequential:
                            color = colors[i % colors.Count];
                            break;
                        case ColorApplyMode.Random:
                            color = colors[random.Next(colors.Count)];
                            break;
                        default:
                            color = colors[i % colors.Count];
                            break;
                    }

                    // 塗りつぶし色を適用
                    if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = color;
                    }
                }
            }, "図形に色を適用");
        }

        /// <summary>
        /// 既存のカラーパレットを削除
        /// </summary>
        /// <param name="slide">対象スライド</param>
        public void RemoveExistingPalettes(PowerPoint.Slide slide)
        {
            if (slide == null)
                return;

            ComExceptionHandler.ExecuteComOperation(() =>
            {
                var shapesToDelete = new List<PowerPoint.Shape>();

                // スライド領域外の図形を検索
                var presentation = slide.Parent as PowerPoint.Presentation;
                float slideWidth = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    // スライド領域外にある正方形図形を削除対象とする
                    bool isOutsideSlide = shape.Left >= slideWidth || shape.Top >= slideHeight;
                    bool isSquare = Math.Abs(shape.Width - shape.Height) < 1f;
                    bool isSmall = shape.Width <= 50f; // パレット用のセルサイズ範囲

                    if (isOutsideSlide && isSquare && isSmall)
                    {
                        shapesToDelete.Add(shape);
                    }
                }

                // 削除実行
                foreach (var shape in shapesToDelete)
                {
                    shape.Delete();
                }
            }, "既存パレット削除", suppressErrors: true);
        }
    }

    /// <summary>
    /// 色の適用モード
    /// </summary>
    public enum ColorApplyMode
    {
        /// <summary>順番に適用</summary>
        Sequential,
        /// <summary>ランダムに適用</summary>
        Random
    }
}
