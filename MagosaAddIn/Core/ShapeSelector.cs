using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形選択補助機能を提供するクラス
    /// </summary>
    public class ShapeSelector
    {
        /// <summary>
        /// 基準図形と同一書式の図形を選択状態にする
        /// </summary>
        /// <param name="baseShape">基準図形</param>
        /// <param name="criteria">選択条件</param>
        public void SelectShapesByFormat(PowerPoint.Shape baseShape, SelectionCriteria criteria)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    if (baseShape == null)
                    {
                        throw new ArgumentNullException(nameof(baseShape), "基準図形が指定されていません。");
                    }

                    // 基準図形のスライドを取得
                    var slide = baseShape.Parent as PowerPoint.Slide;
                    if (slide == null)
                    {
                        throw new InvalidOperationException("基準図形が有効なスライドに配置されていません。");
                    }

                    // 基準図形の書式情報を取得
                    var baseFormat = ExtractShapeFormat(baseShape);
                    if (baseFormat == null)
                    {
                        throw new InvalidOperationException("基準図形の書式情報を取得できませんでした。");
                    }

                    // 同一書式の図形を検索（基準図形も含む）
                    var matchingShapes = FindMatchingShapes(slide, baseShape, baseFormat, criteria);

                    if (matchingShapes.Count == 0)
                    {
                        ComExceptionHandler.LogWarning($"同一書式の図形が見つかりませんでした。条件: {criteria}");
                        return;
                    }

                    // 図形を選択状態にする
                    SelectShapes(slide, matchingShapes);

                    ComExceptionHandler.LogDebug($"同一書式選択完了: {matchingShapes.Count}個の図形を選択, 条件: {criteria}");
                },
                "同一書式図形選択");
        }

        /// <summary>
        /// 同一書式の図形を検索する（基準図形も含む）
        /// </summary>
        /// <param name="slide">検索対象スライド</param>
        /// <param name="baseShape">基準図形</param>
        /// <param name="baseFormat">基準書式</param>
        /// <param name="criteria">選択条件</param>
        /// <returns>一致する図形のリスト（基準図形も含む）</returns>
        private List<PowerPoint.Shape> FindMatchingShapes(PowerPoint.Slide slide, PowerPoint.Shape baseShape, ShapeFormatInfo baseFormat, SelectionCriteria criteria)
        {
            var matchingShapes = new List<PowerPoint.Shape>();

            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    ComExceptionHandler.LogDebug($"図形検索開始 - 条件: {criteria}, 基準図形: {baseFormat.ShapeName}");
                    ComExceptionHandler.LogDebug($"基準書式 - {baseFormat}");

                    // まず基準図形を必ず追加
                    matchingShapes.Add(baseShape);
                    ComExceptionHandler.LogDebug($"基準図形を追加: {baseFormat.ShapeName}");

                    // 基準図形のID/名前を取得（重複チェック用）
                    string baseShapeName = ComExceptionHandler.ExecuteComOperation(
                        () => baseShape.Name,
                        "基準図形名取得",
                        defaultValue: "",
                        suppressErrors: true);

                    int baseShapeId = ComExceptionHandler.ExecuteComOperation(
                        () => baseShape.Id,
                        "基準図形ID取得",
                        defaultValue: -1,
                        suppressErrors: true);

                    ComExceptionHandler.LogDebug($"スライド内図形数: {slide.Shapes.Count}");

                    // スライド内の全図形を検索
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];

                        // 基準図形と同じ図形はスキップ（既に追加済み）
                        int currentShapeId = ComExceptionHandler.ExecuteComOperation(
                            () => shape.Id,
                            "図形ID取得",
                            defaultValue: -2,
                            suppressErrors: true);

                        if (currentShapeId == baseShapeId)
                        {
                            ComExceptionHandler.LogDebug($"基準図形をスキップ: ID={currentShapeId}");
                            continue;
                        }

                        // 図形の書式を取得して比較
                        var shapeFormat = ExtractShapeFormat(shape);
                        if (shapeFormat != null)
                        {
                            ComExceptionHandler.LogDebug($"図形{i}の書式 - {shapeFormat}");

                            if (IsFormatMatch(baseFormat, shapeFormat, criteria))
                            {
                                matchingShapes.Add(shape);

                                string shapeName = ComExceptionHandler.ExecuteComOperation(
                                    () => shape.Name,
                                    "図形名取得",
                                    defaultValue: $"図形{i}",
                                    suppressErrors: true);

                                ComExceptionHandler.LogDebug($"★一致図形追加: {shapeName} (ID={currentShapeId})");
                            }
                            else
                            {
                                ComExceptionHandler.LogDebug($"図形{i}は条件に不一致");
                            }
                        }
                        else
                        {
                            ComExceptionHandler.LogWarning($"図形{i}の書式取得に失敗");
                        }
                    }

                    ComExceptionHandler.LogDebug($"検索完了: {matchingShapes.Count}個の図形が条件に一致");
                    return matchingShapes;
                },
                "同一書式図形検索",
                defaultValue: matchingShapes,
                suppressErrors: true);
        }

        /// <summary>
        /// 図形の書式情報を抽出する
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>書式情報</returns>
        private ShapeFormatInfo ExtractShapeFormat(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var format = new ShapeFormatInfo
                    {
                        ShapeName = shape.Name
                    };

                    // シェイプ種類情報を取得
                    try
                    {
                        format.ShapeType = shape.Type;
                        if (shape.Type == Office.MsoShapeType.msoAutoShape)
                        {
                            format.AutoShapeType = shape.AutoShapeType;
                        }
                        else
                        {
                            format.AutoShapeType = Office.MsoAutoShapeType.msoShapeMixed;
                        }
                    }
                    catch
                    {
                        format.ShapeType = Office.MsoShapeType.msoShapeTypeMixed;
                        format.AutoShapeType = Office.MsoAutoShapeType.msoShapeMixed;
                        ComExceptionHandler.LogWarning($"図形 {shape.Name} の種類情報取得に失敗");
                    }

                    // 塗りつぶし情報を取得
                    try
                    {
                        if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                        {
                            format.HasFill = true;
                            format.FillColor = shape.Fill.ForeColor.RGB;
                            format.FillTransparency = shape.Fill.Transparency;
                        }
                        else
                        {
                            format.HasFill = false;
                        }
                    }
                    catch
                    {
                        format.HasFill = false;
                        ComExceptionHandler.LogWarning($"図形 {shape.Name} の塗りつぶし情報取得に失敗");
                    }

                    // 枠線情報を取得
                    try
                    {
                        if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            format.HasLine = true;
                            format.LineColor = shape.Line.ForeColor.RGB;
                            format.LineWeight = shape.Line.Weight;
                            format.LineDashStyle = shape.Line.DashStyle;
                        }
                        else
                        {
                            format.HasLine = false;
                        }
                    }
                    catch
                    {
                        format.HasLine = false;
                        ComExceptionHandler.LogWarning($"図形 {shape.Name} の枠線情報取得に失敗");
                    }

                    return format;
                },
                $"図形書式抽出: {shape.Name}",
                defaultValue: null,
                suppressErrors: true);
        }

        /// <summary>
        /// 書式が一致するかどうかを判定する
        /// </summary>
        /// <param name="baseFormat">基準書式</param>
        /// <param name="targetFormat">比較対象書式</param>
        /// <param name="criteria">比較条件</param>
        /// <returns>一致する場合true</returns>
        private bool IsFormatMatch(ShapeFormatInfo baseFormat, ShapeFormatInfo targetFormat, SelectionCriteria criteria)
        {
            switch (criteria)
            {
                case SelectionCriteria.FillColorOnly:
                    return IsFillColorMatch(baseFormat, targetFormat);

                case SelectionCriteria.LineStyleOnly:
                    return IsLineStyleMatch(baseFormat, targetFormat);

                case SelectionCriteria.FillAndLineStyle:
                    return IsFillColorMatch(baseFormat, targetFormat) && IsLineStyleMatch(baseFormat, targetFormat);

                case SelectionCriteria.ShapeTypeOnly:
                    return IsShapeTypeMatch(baseFormat, targetFormat);

                default:
                    return false;
            }
        }

        /// <summary>
        /// 塗りつぶし色が一致するかどうかを判定する
        /// </summary>
        private bool IsFillColorMatch(ShapeFormatInfo baseFormat, ShapeFormatInfo targetFormat)
        {
            // 両方とも塗りつぶしなしの場合
            if (!baseFormat.HasFill && !targetFormat.HasFill)
                return true;

            // 片方だけ塗りつぶしなしの場合
            if (baseFormat.HasFill != targetFormat.HasFill)
                return false;

            // 両方とも塗りつぶしありの場合、色を比較
            return baseFormat.FillColor == targetFormat.FillColor;
        }

        /// <summary>
        /// 枠線スタイルが一致するかどうかを判定する
        /// </summary>
        private bool IsLineStyleMatch(ShapeFormatInfo baseFormat, ShapeFormatInfo targetFormat)
        {
            // 両方とも枠線なしの場合
            if (!baseFormat.HasLine && !targetFormat.HasLine)
                return true;

            // 片方だけ枠線なしの場合
            if (baseFormat.HasLine != targetFormat.HasLine)
                return false;

            // 両方とも枠線ありの場合、スタイルを比較
            return baseFormat.LineColor == targetFormat.LineColor &&
                   Math.Abs(baseFormat.LineWeight - targetFormat.LineWeight) < 0.1f &&
                   baseFormat.LineDashStyle == targetFormat.LineDashStyle;
        }

        /// <summary>
        /// シェイプ種類が一致するかどうかを判定する
        /// </summary>
        private bool IsShapeTypeMatch(ShapeFormatInfo baseFormat, ShapeFormatInfo targetFormat)
        {
            // 基本的なシェイプタイプが異なる場合
            if (baseFormat.ShapeType != targetFormat.ShapeType)
                return false;

            // オートシェイプの場合は、オートシェイプタイプも比較
            if (baseFormat.ShapeType == Office.MsoShapeType.msoAutoShape)
            {
                return baseFormat.AutoShapeType == targetFormat.AutoShapeType;
            }

            // その他のシェイプタイプの場合は、基本タイプが一致すればOK
            return true;
        }

        /// <summary>
        /// 図形を選択状態にする
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="shapes">選択する図形のリスト</param>
        private void SelectShapes(PowerPoint.Slide slide, List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    if (shapes.Count == 0)
                    {
                        ComExceptionHandler.LogWarning("選択する図形がありません");
                        return;
                    }

                    // PowerPointアプリケーションを取得
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWindow == null)
                    {
                        throw new InvalidOperationException("PowerPointのアクティブウィンドウが見つかりません。");
                    }

                    ComExceptionHandler.LogDebug($"図形選択開始: {shapes.Count}個の図形を選択");

                    // 方法1: ShapeRangeを使用した一括選択を試行
                    try
                    {
                        // 図形名の配列を作成
                        var shapeNames = new string[shapes.Count];
                        for (int i = 0; i < shapes.Count; i++)
                        {
                            shapeNames[i] = shapes[i].Name;
                            ComExceptionHandler.LogDebug($"選択対象図形: {shapeNames[i]}");
                        }

                        // ShapeRangeで一括選択
                        var shapeRange = slide.Shapes.Range(shapeNames);
                        shapeRange.Select(Office.MsoTriState.msoFalse);

                        ComExceptionHandler.LogDebug($"ShapeRangeによる一括選択成功: {shapes.Count}個");

                        // 選択結果を確認
                        var selection = app.ActiveWindow.Selection;
                        if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            ComExceptionHandler.LogDebug($"選択確認: {selection.ShapeRange.Count}個の図形が選択状態");
                            if (selection.ShapeRange.Count == shapes.Count)
                            {
                                return; // 成功
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ComExceptionHandler.LogWarning($"ShapeRangeによる一括選択に失敗: {ex.Message}");
                    }

                    // 方法2: 個別選択（フォールバック）
                    ComExceptionHandler.LogDebug("個別選択方式にフォールバック");

                    bool firstSelection = true;
                    int successCount = 0;

                    foreach (var shape in shapes)
                    {
                        try
                        {
                            if (firstSelection)
                            {
                                // 最初の図形は既存選択をクリア
                                shape.Select(Office.MsoTriState.msoFalse);
                                firstSelection = false;
                                ComExceptionHandler.LogDebug($"最初の図形を選択: {shape.Name}");
                            }
                            else
                            {
                                // 2番目以降は追加選択
                                shape.Select(Office.MsoTriState.msoTrue);
                                ComExceptionHandler.LogDebug($"図形を追加選択: {shape.Name}");
                            }
                            successCount++;

                            // 少し待機（PowerPointの処理を安定させる）
                            System.Threading.Thread.Sleep(10);
                        }
                        catch (Exception ex)
                        {
                            ComExceptionHandler.LogError($"図形選択に失敗: {shape.Name}", ex);
                        }
                    }

                    // 最終的な選択結果を確認
                    try
                    {
                        var finalSelection = app.ActiveWindow.Selection;
                        if (finalSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            ComExceptionHandler.LogDebug($"最終選択結果: {finalSelection.ShapeRange.Count}個の図形が選択状態（成功: {successCount}/{shapes.Count}）");
                        }
                    }
                    catch (Exception ex)
                    {
                        ComExceptionHandler.LogWarning($"最終選択結果の確認に失敗: {ex.Message}");
                    }
                },
                "図形選択実行");
        }

        /// <summary>
        /// 選択可能な図形数を取得する（プレビュー用）
        /// </summary>
        /// <param name="baseShape">基準図形</param>
        /// <param name="criteria">選択条件</param>
        /// <returns>選択可能な図形数（基準図形も含む）</returns>
        public int GetMatchingShapeCount(PowerPoint.Shape baseShape, SelectionCriteria criteria)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    if (baseShape == null)
                        return 0;

                    var slide = baseShape.Parent as PowerPoint.Slide;
                    if (slide == null)
                        return 0;

                    var baseFormat = ExtractShapeFormat(baseShape);
                    if (baseFormat == null)
                        return 0;

                    var matchingShapes = FindMatchingShapes(slide, baseShape, baseFormat, criteria);
                    return matchingShapes.Count;
                },
                "一致図形数取得",
                defaultValue: 0,
                suppressErrors: true);
        }
    }

    /// <summary>
    /// 選択条件の列挙型
    /// </summary>
    public enum SelectionCriteria
    {
        /// <summary>
        /// 塗りのカラーコードが同じもの
        /// </summary>
        FillColorOnly,

        /// <summary>
        /// 枠線のスタイルが同じもの
        /// </summary>
        LineStyleOnly,

        /// <summary>
        /// 塗りと枠線のスタイルが同じもの
        /// </summary>
        FillAndLineStyle,

        /// <summary>
        /// シェイプの種類が同じもの
        /// </summary>
        ShapeTypeOnly
    }

    /// <summary>
    /// 図形の書式情報を格納するクラス
    /// </summary>
    public class ShapeFormatInfo
    {
        public string ShapeName { get; set; }

        // 塗りつぶし情報
        public bool HasFill { get; set; }
        public int FillColor { get; set; }
        public float FillTransparency { get; set; }

        // 枠線情報
        public bool HasLine { get; set; }
        public int LineColor { get; set; }
        public float LineWeight { get; set; }
        public Office.MsoLineDashStyle LineDashStyle { get; set; }

        // シェイプ種類情報
        public Office.MsoShapeType ShapeType { get; set; }
        public Office.MsoAutoShapeType AutoShapeType { get; set; }

        /// <summary>
        /// 書式情報の文字列表現
        /// </summary>
        public override string ToString()
        {
            var fillInfo = HasFill ? $"塗り: #{FillColor:X6}" : "塗り: なし";
            var lineInfo = HasLine ? $"枠線: #{LineColor:X6}, {LineWeight}pt, {LineDashStyle}" : "枠線: なし";
            var shapeInfo = $"種類: {GetShapeTypeDisplayName()}";
            return $"{ShapeName} - {fillInfo}, {lineInfo}, {shapeInfo}";
        }

        /// <summary>
        /// シェイプ種類の表示名を取得
        /// </summary>
        public string GetShapeTypeDisplayName()
        {
            switch (ShapeType)
            {
                case Office.MsoShapeType.msoAutoShape:
                    switch (AutoShapeType)
                    {
                        case Office.MsoAutoShapeType.msoShapeRectangle:
                            return "四角形";
                        case Office.MsoAutoShapeType.msoShapeRoundedRectangle:
                            return "角丸四角形";
                        case Office.MsoAutoShapeType.msoShapeOval:
                            return "楕円";
                        case Office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                            return "三角形";
                        case Office.MsoAutoShapeType.msoShapeDiamond:
                            return "ひし形";
                        case Office.MsoAutoShapeType.msoShapeHexagon:
                            return "六角形";
                        case Office.MsoAutoShapeType.msoShapeOctagon:
                            return "八角形";
                        default:
                            return $"オートシェイプ({AutoShapeType})";
                    }
                case Office.MsoShapeType.msoTextBox:
                    return "テキストボックス";
                case Office.MsoShapeType.msoPicture:
                    return "画像";
                case Office.MsoShapeType.msoLine:
                    return "線";
                case Office.MsoShapeType.msoFreeform:
                    return "フリーフォーム";
                case Office.MsoShapeType.msoGroup:
                    return "グループ";
                case Office.MsoShapeType.msoTable:
                    return "表";
                case Office.MsoShapeType.msoChart:
                    return "グラフ";
                default:
                    return ShapeType.ToString();
            }
        }
    }
}