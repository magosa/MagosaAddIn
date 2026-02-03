using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形置き換え機能を提供するクラス
    /// </summary>
    public class ShapeReplacer
    {
        private List<ShapeInfo> savedShapes;

        public ShapeReplacer()
        {
            savedShapes = new List<ShapeInfo>();
        }

        /// <summary>
        /// 複数図形を記憶（選択完了時）
        /// </summary>
        /// <param name="shapes">記憶する図形リスト</param>
        public void SaveShapes(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_REPLACEMENT, "図形記憶");

                    // 既存の記憶をクリア
                    savedShapes.Clear();

                    // 図形情報を事前取得して保存
                    foreach (var shape in shapes)
                    {
                        var info = ExtractShapeInfo(shape);
                        if (info != null)
                        {
                            savedShapes.Add(info);
                            ComExceptionHandler.LogDebug($"図形記憶: {info.ShapeName} ({info.Left:F1}, {info.Top:F1}, {info.Width:F1}×{info.Height:F1})");
                        }
                    }

                    ComExceptionHandler.LogDebug($"図形記憶完了: {savedShapes.Count}個");
                },
                "図形記憶");
        }

        /// <summary>
        /// テンプレート図形で一括置き換え
        /// </summary>
        /// <param name="templateShape">テンプレート図形</param>
        /// <param name="options">置き換えオプション</param>
        public void ReplaceShapes(PowerPoint.Shape templateShape, ReplacementOptions options)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (templateShape == null)
                    {
                        throw new ArgumentNullException(nameof(templateShape), "テンプレート図形が指定されていません。");
                    }

                    if (savedShapes == null || savedShapes.Count == 0)
                    {
                        throw new InvalidOperationException("置き換え対象の図形が記憶されていません。先に「選択完了」を実行してください。");
                    }

                    var slide = ComExceptionHandler.ExecuteComOperation(
                        () => templateShape.Parent as PowerPoint.Slide,
                        "スライド取得");

                    if (slide == null)
                    {
                        throw new InvalidOperationException("テンプレート図形が有効なスライドに配置されていません。");
                    }

                    // テンプレート図形の情報を取得
                    var templateInfo = ExtractShapeInfo(templateShape);
                    if (templateInfo == null)
                    {
                        throw new InvalidOperationException("テンプレート図形の情報を取得できませんでした。");
                    }

                    ComExceptionHandler.LogDebug($"=== 図形置き換え開始 ===");
                    ComExceptionHandler.LogDebug($"テンプレート: {templateInfo.ShapeName}");
                    ComExceptionHandler.LogDebug($"対象図形数: {savedShapes.Count}");
                    ComExceptionHandler.LogDebug($"サイズモード: {options.SizeMode}");
                    ComExceptionHandler.LogDebug($"スタイル継承: {options.InheritStyle}");
                    ComExceptionHandler.LogDebug($"テキスト継承: {options.InheritText}");

                    int successCount = 0;
                    var createdShapes = new List<PowerPoint.Shape>();

                    // 各保存図形を置き換え
                    foreach (var savedShape in savedShapes)
                    {
                        try
                        {
                            // テンプレート図形を複製
                            var newShape = ComExceptionHandler.ExecuteComOperation(
                                () => templateShape.Duplicate()[1],
                                $"図形複製: {savedShape.ShapeName}");

                            if (newShape != null)
                            {
                                // サイズを設定
                                SetShapeSize(newShape, savedShape, templateInfo, options.SizeMode);

                                // 中心点で位置を設定
                                SetShapePosition(newShape, savedShape);

                                // スタイルを適用
                                if (options.InheritStyle)
                                {
                                    ApplyShapeStyle(newShape, savedShape);
                                }

                                // テキストを適用
                                if (options.InheritText && !string.IsNullOrEmpty(savedShape.Text))
                                {
                                    ApplyShapeText(newShape, savedShape.Text);
                                }

                                createdShapes.Add(newShape);
                                successCount++;

                                ComExceptionHandler.LogDebug($"置き換え成功: {savedShape.ShapeName} → 新規図形");
                            }
                        }
                        catch (Exception ex)
                        {
                            ComExceptionHandler.LogError($"図形置き換えエラー: {savedShape.ShapeName}", ex);
                        }
                    }

                    // 元の図形を削除
                    DeleteOriginalShapes();

                    ComExceptionHandler.LogDebug($"=== 図形置き換え完了: {successCount}/{savedShapes.Count}個 ===");

                    // 記憶をクリア
                    ClearSavedShapes();
                },
                "図形一括置き換え");
        }

        /// <summary>
        /// 記憶した図形数を取得
        /// </summary>
        /// <returns>記憶した図形数</returns>
        public int GetSavedShapeCount()
        {
            return savedShapes?.Count ?? 0;
        }

        /// <summary>
        /// 記憶した図形をクリア
        /// </summary>
        public void ClearSavedShapes()
        {
            savedShapes?.Clear();
            ComExceptionHandler.LogDebug("記憶図形をクリアしました");
        }

        #region プライベートメソッド

        /// <summary>
        /// 図形情報を抽出
        /// </summary>
        private ShapeInfo ExtractShapeInfo(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    var info = new ShapeInfo
                    {
                        ShapeName = shape.Name,
                        Left = shape.Left,
                        Top = shape.Top,
                        Width = shape.Width,
                        Height = shape.Height,
                        OriginalShape = shape
                    };

                    // 中心点を計算
                    info.CenterX = info.Left + (info.Width / 2);
                    info.CenterY = info.Top + (info.Height / 2);

                    // スタイル情報を取得
                    try
                    {
                        if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                        {
                            info.FillColor = shape.Fill.ForeColor.RGB;
                            info.FillTransparency = shape.Fill.Transparency;
                        }

                        if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            info.LineColor = shape.Line.ForeColor.RGB;
                            info.LineWeight = shape.Line.Weight;
                            info.LineDashStyle = shape.Line.DashStyle;
                        }

                        if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                        {
                            info.HasShadow = true;
                            info.ShadowColor = shape.Shadow.ForeColor.RGB;
                        }
                    }
                    catch
                    {
                        // スタイル取得失敗時はデフォルト値を使用
                        ComExceptionHandler.LogWarning($"図形 {info.ShapeName} のスタイル情報取得に失敗");
                    }

                    // テキスト情報を取得
                    try
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                            shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            info.Text = shape.TextFrame.TextRange.Text;
                        }
                    }
                    catch
                    {
                        ComExceptionHandler.LogWarning($"図形 {info.ShapeName} のテキスト情報取得に失敗");
                    }

                    return info;
                },
                $"図形情報抽出: {shape.Name}",
                suppressErrors: true);
        }

        /// <summary>
        /// 図形のサイズを設定
        /// </summary>
        private void SetShapeSize(PowerPoint.Shape newShape, ShapeInfo savedShape, ShapeInfo templateInfo, SizeMode sizeMode)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (sizeMode == SizeMode.KeepOriginal)
                    {
                        // 元のサイズを維持
                        newShape.Width = savedShape.Width;
                        newShape.Height = savedShape.Height;
                    }
                    // SizeMode.UseTemplateの場合はテンプレートのサイズをそのまま使用（何もしない）
                },
                "サイズ設定",
                suppressErrors: true);
        }

        /// <summary>
        /// 図形の位置を中心点基準で設定
        /// </summary>
        private void SetShapePosition(PowerPoint.Shape newShape, ShapeInfo savedShape)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    // 新しい図形の中心点が元の図形の中心点と一致するように配置
                    float newLeft = savedShape.CenterX - (newShape.Width / 2);
                    float newTop = savedShape.CenterY - (newShape.Height / 2);

                    // 座標の検証
                    if (ErrorHandler.ValidateCoordinates(newLeft, newTop, "図形配置"))
                    {
                        newShape.Left = newLeft;
                        newShape.Top = newTop;
                    }
                },
                "位置設定",
                suppressErrors: true);
        }

        /// <summary>
        /// 図形にスタイルを適用
        /// </summary>
        private void ApplyShapeStyle(PowerPoint.Shape newShape, ShapeInfo savedShape)
        {
            // 塗りつぶしを適用
            if (savedShape.FillColor.HasValue)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        newShape.Fill.Visible = Office.MsoTriState.msoTrue;
                        newShape.Fill.ForeColor.RGB = savedShape.FillColor.Value;
                        newShape.Fill.Transparency = savedShape.FillTransparency;
                    },
                    "塗りつぶし適用",
                    suppressErrors: true);
            }

            // 枠線を適用
            if (savedShape.LineColor.HasValue)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        newShape.Line.Visible = Office.MsoTriState.msoTrue;
                        newShape.Line.ForeColor.RGB = savedShape.LineColor.Value;
                        newShape.Line.Weight = savedShape.LineWeight;
                        newShape.Line.DashStyle = savedShape.LineDashStyle;
                    },
                    "枠線適用",
                    suppressErrors: true);
            }

            // 影を適用
            if (savedShape.HasShadow)
            {
                ComExceptionHandler.ExecuteComOperation(
                    () =>
                    {
                        newShape.Shadow.Visible = Office.MsoTriState.msoTrue;
                        newShape.Shadow.ForeColor.RGB = savedShape.ShadowColor;
                    },
                    "影適用",
                    suppressErrors: true);
            }
        }

        /// <summary>
        /// 図形にテキストを適用
        /// </summary>
        private void ApplyShapeText(PowerPoint.Shape newShape, string text)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (newShape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        newShape.TextFrame.TextRange.Text = text;
                    }
                },
                "テキスト適用",
                suppressErrors: true);
        }

        /// <summary>
        /// 元の図形を削除
        /// </summary>
        private void DeleteOriginalShapes()
        {
            int deletedCount = 0;
            foreach (var savedShape in savedShapes)
            {
                if (savedShape.OriginalShape != null)
                {
                    try
                    {
                        ComExceptionHandler.ExecuteComOperation(
                            () => savedShape.OriginalShape.Delete(),
                            $"図形削除: {savedShape.ShapeName}",
                            suppressErrors: true);
                        deletedCount++;
                    }
                    catch
                    {
                        ComExceptionHandler.LogWarning($"図形削除失敗: {savedShape.ShapeName}");
                    }
                }
            }
            ComExceptionHandler.LogDebug($"元図形削除完了: {deletedCount}/{savedShapes.Count}個");
        }

        #endregion
    }

    /// <summary>
    /// 置き換えオプション
    /// </summary>
    public class ReplacementOptions
    {
        /// <summary>
        /// サイズモード
        /// </summary>
        public SizeMode SizeMode { get; set; }

        /// <summary>
        /// スタイルを継承するか（塗りつぶし・枠線・影）
        /// </summary>
        public bool InheritStyle { get; set; }

        /// <summary>
        /// テキストを継承するか
        /// </summary>
        public bool InheritText { get; set; }

        public ReplacementOptions()
        {
            SizeMode = SizeMode.KeepOriginal;
            InheritStyle = false;
            InheritText = false;
        }
    }

    /// <summary>
    /// サイズモード
    /// </summary>
    public enum SizeMode
    {
        /// <summary>
        /// 元のサイズを維持
        /// </summary>
        KeepOriginal,

        /// <summary>
        /// テンプレートサイズに統一
        /// </summary>
        UseTemplate
    }
}