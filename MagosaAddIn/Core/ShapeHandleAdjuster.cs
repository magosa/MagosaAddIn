using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形ハンドル調整機能を提供するクラス
    /// </summary>
    public class ShapeHandleAdjuster
    {
        #region 調整ハンドル設定

        /// <summary>
        /// 複数の調整ハンドルを一括設定（正規化値）
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="handleValues">ハンドル値の配列（インデックス0 = ハンドル1）</param>
        public void AdjustHandles(List<PowerPoint.Shape> shapes, float[] handleValues)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "調整ハンドル設定");

                    if (handleValues == null || handleValues.Length == 0)
                    {
                        throw new ArgumentException("ハンドル値が指定されていません");
                    }

                    // 各ハンドル値を検証
                    for (int i = 0; i < handleValues.Length; i++)
                    {
                        ErrorHandler.ValidateRange(handleValues[i], Constants.MIN_HANDLE_VALUE,
                            Constants.MAX_HANDLE_VALUE, $"ハンドル{i + 1}", "調整ハンドル設定");
                    }

                    int successCount = 0;
                    foreach (var shape in shapes)
                    {
                        if (SetMultipleAdjustmentHandles(shape, handleValues))
                            successCount++;
                    }

                    ComExceptionHandler.LogDebug($"調整ハンドル設定完了: 成功 {successCount}/{shapes.Count}, " +
                        $"ハンドル数: {handleValues.Length}");
                },
                "調整ハンドル設定");
        }

        /// <summary>
        /// mm単位で調整ハンドルを設定
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="handleValuesInMm">ハンドル値の配列（mm単位）</param>
        public void AdjustHandlesInMm(List<PowerPoint.Shape> shapes, float[] handleValuesInMm)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "調整ハンドル設定（mm）");

                    if (handleValuesInMm == null || handleValuesInMm.Length == 0)
                    {
                        throw new ArgumentException("ハンドル値が指定されていません");
                    }

                    // 各ハンドル値を検証
                    for (int i = 0; i < handleValuesInMm.Length; i++)
                    {
                        ErrorHandler.ValidateRange(handleValuesInMm[i], Constants.MIN_HANDLE_MM,
                            Constants.MAX_HANDLE_MM, $"ハンドル{i + 1}（mm）", "調整ハンドル設定（mm）");
                    }

                    int successCount = 0;
                    foreach (var shape in shapes)
                    {
                        // mm値を正規化値に変換してから設定
                        var normalizedValues = new float[handleValuesInMm.Length];
                        for (int i = 0; i < handleValuesInMm.Length; i++)
                        {
                            normalizedValues[i] = ConvertMmToNormalized(handleValuesInMm[i], shape, i);
                        }

                        if (SetMultipleAdjustmentHandles(shape, normalizedValues))
                            successCount++;
                    }

                    ComExceptionHandler.LogDebug($"調整ハンドル設定（mm）完了: 成功 {successCount}/{shapes.Count}, " +
                        $"ハンドル数: {handleValuesInMm.Length}");
                },
                "調整ハンドル設定（mm）");
        }

        /// <summary>
        /// 度数単位で角度ハンドルを設定
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <param name="angleValuesInDegree">角度値の配列（度数単位）</param>
        public void AdjustAngleHandlesInDegree(List<PowerPoint.Shape> shapes, float[] angleValuesInDegree)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "角度ハンドル設定（度）");

                    if (angleValuesInDegree == null || angleValuesInDegree.Length == 0)
                    {
                        throw new ArgumentException("角度値が指定されていません");
                    }

                    // 各角度値を検証
                    for (int i = 0; i < angleValuesInDegree.Length; i++)
                    {
                        ErrorHandler.ValidateRange(angleValuesInDegree[i], Constants.MIN_ANGLE_DEGREE,
                            Constants.MAX_ANGLE_DEGREE, $"角度{i + 1}（度）", "角度ハンドル設定（度）");
                    }

                    int successCount = 0;
                    foreach (var shape in shapes)
                    {
                        // 図形タイプを取得
                        var shapeType = shape.AutoShapeType;

                        // 設定前の現在値をログ出力
                        ComExceptionHandler.LogDebug($"=== 角度設定前の状態 ===");
                        ComExceptionHandler.LogDebug($"図形: {shape.Name}, タイプ: {shapeType}");
                        for (int i = 1; i <= shape.Adjustments.Count; i++)
                        {
                            ComExceptionHandler.LogDebug($"設定前ハンドル{i}: {shape.Adjustments[i]}");
                        }

                        // 度数値を正規化値に変換してから設定（図形タイプ別）
                        var normalizedValues = new float[angleValuesInDegree.Length];
                        for (int i = 0; i < angleValuesInDegree.Length; i++)
                        {
                            normalizedValues[i] = ConvertDegreeToNormalizedByShapeType(angleValuesInDegree[i], shapeType, i);
                            ComExceptionHandler.LogDebug($"角度変換: {angleValuesInDegree[i]}° → {normalizedValues[i]:F3} (図形: {shapeType}, ハンドル{i + 1})");
                        }

                        if (SetMultipleAdjustmentHandles(shape, normalizedValues))
                        {
                            // 設定後の値をログ出力
                            ComExceptionHandler.LogDebug($"=== 角度設定後の状態 ===");
                            for (int i = 1; i <= shape.Adjustments.Count; i++)
                            {
                                ComExceptionHandler.LogDebug($"設定後ハンドル{i}: {shape.Adjustments[i]}");
                            }
                            successCount++;
                        }
                    }

                    ComExceptionHandler.LogDebug($"角度ハンドル設定（度）完了: 成功 {successCount}/{shapes.Count}, " +
                        $"角度数: {angleValuesInDegree.Length}");
                },
                "角度ハンドル設定（度）");
        }

        #endregion

        #region リセット機能

        /// <summary>
        /// 図形の調整をリセット
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        public void ResetAdjustments(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_HANDLE_ADJUSTMENT, "調整リセット");

                    int successCount = 0;
                    foreach (var shape in shapes)
                    {
                        if (ResetShapeAdjustments(shape))
                            successCount++;
                    }

                    ComExceptionHandler.LogDebug($"調整リセット完了: 成功 {successCount}/{shapes.Count}");
                },
                "調整リセット");
        }

        #endregion

        #region 単位変換メソッド

        /// <summary>
        /// mm値を正規化値（0.0-1.0）に変換
        /// </summary>
        /// <param name="mmValue">mm値</param>
        /// <param name="shape">対象図形</param>
        /// <param name="handleIndex">ハンドルインデックス</param>
        /// <returns>正規化値</returns>
        private float ConvertMmToNormalized(float mmValue, PowerPoint.Shape shape, int handleIndex)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    // 図形のサイズを取得
                    float shapeWidth = shape.Width;
                    float shapeHeight = shape.Height;
                    float maxDimension = Math.Max(shapeWidth, shapeHeight);

                    // mm→pt変換
                    float ptValue = mmValue * Constants.MM_TO_PT;

                    // 図形サイズに対する比率として正規化
                    float normalizedValue = ptValue / maxDimension;

                    // 0.0-1.0の範囲にクランプ
                    return Math.Max(0.0f, Math.Min(1.0f, normalizedValue));
                },
                $"mm→正規化変換: {mmValue}mm",
                defaultValue: Constants.DEFAULT_HANDLE_VALUE,
                suppressErrors: true);
        }

        /// <summary>
        /// 正規化値をmm値に変換
        /// </summary>
        /// <param name="normalizedValue">正規化値</param>
        /// <param name="shape">対象図形</param>
        /// <param name="handleIndex">ハンドルインデックス</param>
        /// <returns>mm値</returns>
        public float ConvertNormalizedToMm(float normalizedValue, PowerPoint.Shape shape, int handleIndex)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    // 図形のサイズを取得
                    float shapeWidth = shape.Width;
                    float shapeHeight = shape.Height;
                    float maxDimension = Math.Max(shapeWidth, shapeHeight);

                    // 正規化値→pt変換
                    float ptValue = normalizedValue * maxDimension;

                    // pt→mm変換
                    return ptValue * Constants.PT_TO_MM;
                },
                $"正規化→mm変換: {normalizedValue}",
                defaultValue: Constants.DEFAULT_HANDLE_MM,
                suppressErrors: true);
        }

        /// <summary>
        /// 度数を正規化値（0.0-1.0）に変換
        /// </summary>
        /// <param name="degreeValue">度数値</param>
        /// <returns>正規化値</returns>
        private float ConvertDegreeToNormalized(float degreeValue)
        {
            // ブロック円弧の場合、角度の範囲が異なる可能性があるため、
            // 0-360度を0.0-1.0に線形変換
            float normalizedDegree = degreeValue % 360.0f;
            if (normalizedDegree < 0) normalizedDegree += 360.0f;

            // 0.0-1.0の範囲に変換
            return normalizedDegree / 360.0f;
        }

        /// <summary>
        /// 正規化値を度数に変換
        /// </summary>
        /// <param name="normalizedValue">正規化値</param>
        /// <returns>度数値</returns>
        private float ConvertNormalizedToDegree(float normalizedValue)
        {
            return normalizedValue * 360.0f;
        }

        /// <summary>
        /// 図形タイプに応じた度数を正規化値に変換（PowerPoint内部形式対応・左回り対応）
        /// </summary>
        /// <param name="degreeValue">度数値（左回り基準）</param>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="handleIndex">ハンドルインデックス</param>
        /// <returns>PowerPoint内部形式の値</returns>
        private float ConvertDegreeToNormalizedByShapeType(float degreeValue, Office.MsoAutoShapeType shapeType, int handleIndex)
        {
            switch (shapeType)
            {
                case Office.MsoAutoShapeType.msoShapePie:
                case Office.MsoAutoShapeType.msoShapeArc:
                case Office.MsoAutoShapeType.msoShapeChord:
                    // 扇形、円弧、弦の場合
                    // 左回り（反時計回り）を右回り（時計回り）に変換
                    float rightRotationDegree = -degreeValue; // 符号を反転

                    // PowerPointでは0-360度を-180～+180度の範囲に変換
                    if (rightRotationDegree > 180)
                    {
                        rightRotationDegree -= 360;
                    }
                    else if (rightRotationDegree < -180)
                    {
                        rightRotationDegree += 360;
                    }

                    ComExceptionHandler.LogDebug($"角度変換（左回り→右回り）: {degreeValue}° → {-degreeValue}° → {rightRotationDegree}°");
                    return rightRotationDegree;

                case Office.MsoAutoShapeType.msoShapeBlockArc:
                    // ブロック円弧の場合
                    if (handleIndex == 0 || handleIndex == 1) // 開始角度、終了角度
                    {
                        float rightRotationBlockArcDegree = -degreeValue; // 符号を反転

                        // 0-360度を-180～+180度に変換
                        if (rightRotationBlockArcDegree > 180)
                        {
                            rightRotationBlockArcDegree -= 360;
                        }
                        else if (rightRotationBlockArcDegree < -180)
                        {
                            rightRotationBlockArcDegree += 360;
                        }

                        ComExceptionHandler.LogDebug($"ブロック円弧角度変換（左回り→右回り）: {degreeValue}° → {rightRotationBlockArcDegree}°");
                        return rightRotationBlockArcDegree;
                    }
                    else if (handleIndex == 2) // 内径比率
                    {
                        // 内径比率は0-100%を0.0-1.0に変換
                        return Math.Max(0.0f, Math.Min(1.0f, degreeValue / 100.0f));
                    }
                    break;

                case Office.MsoAutoShapeType.msoShapeDonut:
                    // ドーナツの場合（内径比率をパーセンテージで入力）
                    return Math.Max(0.0f, Math.Min(1.0f, degreeValue / 100.0f));

                case Office.MsoAutoShapeType.msoShapeMoon:
                    // 三日月の場合
                    float rightRotationMoonDegree = -degreeValue; // 符号を反転

                    // 0-360度を-180～+180度に変換
                    if (rightRotationMoonDegree > 180)
                    {
                        rightRotationMoonDegree -= 360;
                    }
                    else if (rightRotationMoonDegree < -180)
                    {
                        rightRotationMoonDegree += 360;
                    }

                    return rightRotationMoonDegree;

                default:
                    // その他の図形
                    return degreeValue;
            }

            return 0.0f;
        }

        /// <summary>
        /// 図形タイプに応じた正規化値を度数に変換（PowerPoint内部形式対応・左回り対応）
        /// </summary>
        /// <param name="normalizedValue">PowerPoint内部形式の値</param>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="handleIndex">ハンドルインデックス</param>
        /// <returns>度数値（左回り基準）</returns>
        public float ConvertNormalizedToDegreeByShapeType(float normalizedValue, Office.MsoAutoShapeType shapeType, int handleIndex)
        {
            switch (shapeType)
            {
                case Office.MsoAutoShapeType.msoShapePie:
                case Office.MsoAutoShapeType.msoShapeArc:
                case Office.MsoAutoShapeType.msoShapeChord:
                    // 扇形、円弧、弦の場合
                    // PowerPointの内部値（右回り）を左回りに変換
                    float leftRotationDegree = -normalizedValue; // 符号を反転

                    // -180～+180度を0-360度に変換
                    if (leftRotationDegree < 0)
                    {
                        leftRotationDegree += 360;
                    }

                    ComExceptionHandler.LogDebug($"角度表示変換（右回り→左回り）: {normalizedValue}° → {-normalizedValue}° → {leftRotationDegree}°");
                    return leftRotationDegree;

                case Office.MsoAutoShapeType.msoShapeBlockArc:
                    // ブロック円弧の場合
                    if (handleIndex == 0 || handleIndex == 1) // 開始角度、終了角度
                    {
                        float leftRotationBlockArcDegree = -normalizedValue; // 符号を反転

                        // -180～+180度を0-360度に変換
                        if (leftRotationBlockArcDegree < 0)
                        {
                            leftRotationBlockArcDegree += 360;
                        }

                        ComExceptionHandler.LogDebug($"ブロック円弧表示変換（右回り→左回り）: {normalizedValue}° → {leftRotationBlockArcDegree}°");
                        return leftRotationBlockArcDegree;
                    }
                    else if (handleIndex == 2) // 内径比率
                    {
                        return normalizedValue * 100.0f; // パーセンテージで表示
                    }
                    break;

                case Office.MsoAutoShapeType.msoShapeDonut:
                    return normalizedValue * 100.0f; // パーセンテージで表示

                case Office.MsoAutoShapeType.msoShapeMoon:
                    // 三日月の場合
                    float leftRotationMoonDegree = -normalizedValue; // 符号を反転

                    // -180～+180度を0-360度に変換
                    if (leftRotationMoonDegree < 0)
                    {
                        leftRotationMoonDegree += 360;
                    }

                    return leftRotationMoonDegree;

                default:
                    return normalizedValue;
            }

            return 0.0f;
        }

        /// <summary>
        /// PowerPointの角度ハンドル動作を詳細調査
        /// </summary>
        /// <param name="shape">対象図形</param>
        public void InvestigateAngleHandleBehavior(PowerPoint.Shape shape)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    ComExceptionHandler.LogDebug($"=== PowerPoint角度ハンドル調査 ===");
                    ComExceptionHandler.LogDebug($"図形: {shape.Name}, タイプ: {shape.AutoShapeType}");
                    ComExceptionHandler.LogDebug($"ハンドル数: {shape.Adjustments.Count}");

                    // 現在の値を記録
                    var originalValues = new float[shape.Adjustments.Count];
                    for (int i = 1; i <= shape.Adjustments.Count; i++)
                    {
                        originalValues[i - 1] = shape.Adjustments[i];
                        ComExceptionHandler.LogDebug($"元の値 ハンドル{i}: {originalValues[i - 1]}");
                    }

                    // 様々な値でテスト
                    float[] testDegrees = { 0, 30, 45, 90, 135, 180, 270, 360, -30, -90, -180 };

                    ComExceptionHandler.LogDebug($"=== 角度テスト開始 ===");

                    for (int handleIndex = 1; handleIndex <= Math.Min(shape.Adjustments.Count, 2); handleIndex++)
                    {
                        ComExceptionHandler.LogDebug($"--- ハンドル{handleIndex}のテスト ---");

                        foreach (float testDegree in testDegrees)
                        {
                            try
                            {
                                // 元の値に戻す
                                for (int i = 1; i <= shape.Adjustments.Count; i++)
                                {
                                    shape.Adjustments[i] = originalValues[i - 1];
                                }

                                // テスト値を設定
                                shape.Adjustments[handleIndex] = testDegree;
                                float actualValue = shape.Adjustments[handleIndex];

                                ComExceptionHandler.LogDebug($"設定: {testDegree}° → 実際: {actualValue}");

                                // 小さな待機時間
                                System.Threading.Thread.Sleep(10);
                            }
                            catch (Exception ex)
                            {
                                ComExceptionHandler.LogDebug($"エラー: {testDegree}° → {ex.Message}");
                            }
                        }
                    }

                    // 元の値に戻す
                    for (int i = 1; i <= shape.Adjustments.Count; i++)
                    {
                        shape.Adjustments[i] = originalValues[i - 1];
                    }

                    ComExceptionHandler.LogDebug($"=== 調査完了 ===");
                },
                "PowerPoint角度ハンドル調査",
                suppressErrors: true);
        }

        #endregion

        #region 内部メソッド

        /// <summary>
        /// 複数の調整ハンドルを設定
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="handleValues">ハンドル値の配列</param>
        /// <returns>設定成功フラグ</returns>
        private bool SetMultipleAdjustmentHandles(PowerPoint.Shape shape, float[] handleValues)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    int availableHandles = shape.Adjustments.Count;
                    int handlesToSet = Math.Min(handleValues.Length, availableHandles);

                    if (availableHandles == 0)
                    {
                        ComExceptionHandler.LogWarning($"図形 {shape.Name}: 調整ハンドルがありません");
                        return false;
                    }

                    for (int i = 0; i < handlesToSet; i++)
                    {
                        float oldValue = shape.Adjustments[i + 1]; // PowerPointは1ベース
                        shape.Adjustments[i + 1] = handleValues[i];
                        ComExceptionHandler.LogDebug($"ハンドル設定: {shape.Name}[{i + 1}] {oldValue:F3} → {handleValues[i]:F3}");
                    }

                    return true;
                },
                $"調整ハンドル設定: {shape.Name}",
                defaultValue: false,
                suppressErrors: true);
        }

        /// <summary>
        /// 単一図形の調整をリセット
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>リセット成功フラグ</returns>
        private bool ResetShapeAdjustments(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    // 調整ハンドルをデフォルト値にリセット
                    for (int i = 1; i <= shape.Adjustments.Count; i++)
                    {
                        shape.Adjustments[i] = Constants.DEFAULT_HANDLE_VALUE;
                    }

                    ComExceptionHandler.LogDebug($"調整リセット: {shape.Name} (ハンドル数: {shape.Adjustments.Count})");
                    return true;
                },
                $"調整リセット: {shape.Name}",
                defaultValue: false,
                suppressErrors: true);
        }

        #endregion

        #region 情報取得・分析

        /// <summary>
        /// 選択図形群の調整ハンドル情報を分析（高速化版）
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        /// <returns>分析結果</returns>
        public ShapeHandleAnalysis AnalyzeShapes(List<PowerPoint.Shape> shapes)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var analysis = new ShapeHandleAnalysis();

                    if (shapes == null || shapes.Count == 0)
                    {
                        ComExceptionHandler.LogDebug("AnalyzeShapes: 図形リストが空です");
                        return analysis;
                    }
                    ComExceptionHandler.LogDebug($"AnalyzeShapes: {shapes.Count}個の図形を分析開始");

                    foreach (var shape in shapes)
                    {
                        // 高速化: GetHandleInfoFastを使用
                        var info = GetHandleInfoFast(shape);
                        analysis.ShapeInfos.Add(info);

                        ComExceptionHandler.LogDebug($"図形分析: {info.ShapeName} - " +
                            $"タイプ: {info.ShapeType}, " +
                            $"ハンドル数: {info.AdjustmentCount}, " +
                            $"調整可能: {info.IsAdjustmentHandleShape}, " +
                            $"角度ハンドル: {info.IsAngleHandleShape}");

                        // 統計情報を更新
                        analysis.TotalShapes++;
                        analysis.MaxAdjustmentHandles = Math.Max(analysis.MaxAdjustmentHandles, info.AdjustmentCount);

                        if (info.AdjustmentCount > 0)
                        {
                            analysis.ShapesWithAdjustmentHandles++;

                            // 角度ハンドルを持つ図形かどうかを判定
                            if (info.IsAngleHandleShape)
                            {
                                analysis.ShapesWithAngleHandles++;
                                analysis.AngleHandleShapeTypes.Add(info.ShapeType);
                            }
                        }
                    }

                    // 代表的なハンドル数を決定（最頻値）
                    var handleCounts = analysis.ShapeInfos
                        .Where(info => info.AdjustmentCount > 0)
                        .GroupBy(info => info.AdjustmentCount)
                        .OrderByDescending(g => g.Count())
                        .FirstOrDefault();

                    analysis.RecommendedHandleCount = handleCounts?.Key ?? 0;

                    // 角度ハンドル用の推奨ハンドル数を決定
                    var angleHandleCounts = analysis.ShapeInfos
                        .Where(info => info.AdjustmentCount > 0 && info.IsAngleHandleShape)
                        .GroupBy(info => info.AdjustmentCount)
                        .OrderByDescending(g => g.Count())
                        .FirstOrDefault();

                    analysis.RecommendedAngleHandleCount = angleHandleCounts?.Key ?? 0;

                    ComExceptionHandler.LogDebug($"図形分析完了: 総数 {analysis.TotalShapes}, " +
                        $"調整ハンドル有り {analysis.ShapesWithAdjustmentHandles}, " +
                        $"角度ハンドル有り {analysis.ShapesWithAngleHandles}, " +
                        $"推奨ハンドル数 {analysis.RecommendedHandleCount}");

                    return analysis;
                },
                "図形分析",
                defaultValue: new ShapeHandleAnalysis(),
                suppressErrors: true); // 分析処理なのでエラーを抑制
        }

        /// <summary>
        /// 図形の基本ハンドル情報を高速取得（詳細な変換処理は省略）
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>ハンドル情報</returns>
        public ShapeHandleInfo GetHandleInfoFast(PowerPoint.Shape shape) // publicに変更
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var info = new ShapeHandleInfo
                    {
                        ShapeName = shape.Name,
                        ShapeType = shape.AutoShapeType.ToString(),
                        AdjustmentCount = shape.Adjustments.Count,
                        IsAngleHandleShape = IsAngleHandleShape(shape),
                        IsAdjustmentHandleShape = shape.Adjustments.Count > 0, // 簡略化
                        OriginalShape = shape
                    };

                    // 基本的な調整ハンドル値のみ取得（変換処理は後で実行）
                    for (int i = 1; i <= Math.Min(shape.Adjustments.Count, Constants.MAX_SUPPORTED_HANDLES); i++)
                    {
                        info.AdjustmentValues.Add(shape.Adjustments[i]);
                    }

                    // 角度ハンドル図形の場合、角度の意味を解釈（軽量版）
                    if (info.IsAngleHandleShape && info.AdjustmentValues.Count > 0)
                    {
                        info.AngleInterpretation = GetAngleInterpretationFast(shape.AutoShapeType, info.AdjustmentValues.Count);
                    }

                    return info;
                },
                $"ハンドル情報取得（高速）: {shape.Name}",
                defaultValue: new ShapeHandleInfo { ShapeName = "エラー" },
                suppressErrors: true);
        }

        /// <summary>
        /// 図形の現在のハンドル情報を取得（既存メソッドを保持）
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>ハンドル情報</returns>
        public ShapeHandleInfo GetHandleInfo(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var info = new ShapeHandleInfo
                    {
                        ShapeName = shape.Name,
                        ShapeType = shape.AutoShapeType.ToString(),
                        AdjustmentCount = shape.Adjustments.Count,
                        IsAngleHandleShape = IsAngleHandleShape(shape),
                        IsAdjustmentHandleShape = IsAdjustmentHandleShape(shape),
                        OriginalShape = shape
                    };

                    // 調整ハンドル値を取得（最大8個まで）
                    for (int i = 1; i <= Math.Min(shape.Adjustments.Count, Constants.MAX_SUPPORTED_HANDLES); i++)
                    {
                        float normalizedValue = shape.Adjustments[i];
                        info.AdjustmentValues.Add(normalizedValue);

                        // mm値と度数値も計算して保存
                        if (info.IsAngleHandleShape)
                        {
                            info.AdjustmentValuesInDegree.Add(ConvertNormalizedToDegree(normalizedValue));
                        }
                        else
                        {
                            info.AdjustmentValuesInMm.Add(ConvertNormalizedToMm(normalizedValue, shape, i - 1));
                        }
                    }

                    // 角度ハンドル図形の場合、角度の意味を解釈
                    if (info.IsAngleHandleShape && info.AdjustmentValues.Count > 0)
                    {
                        info.AngleInterpretation = GetAngleInterpretation(shape.AutoShapeType, info.AdjustmentValues.Count);
                    }

                    // 調整ハンドル図形の場合、調整の意味を解釈
                    if (info.IsAdjustmentHandleShape && info.AdjustmentValues.Count > 0)
                    {
                        info.AdjustmentInterpretation = GetAdjustmentHandleInterpretation(shape.AutoShapeType, info.AdjustmentValues.Count);
                    }

                    return info;
                },
                $"ハンドル情報取得: {shape.Name}",
                defaultValue: new ShapeHandleInfo { ShapeName = "エラー" },
                suppressErrors: true);
        }

        /// <summary>
        /// 角度ハンドルの意味を高速取得（キャッシュ化）
        /// </summary>
        private static readonly Dictionary<Office.MsoAutoShapeType, List<string>> AngleInterpretationCache =
            new Dictionary<Office.MsoAutoShapeType, List<string>>();

        private List<string> GetAngleInterpretationFast(Office.MsoAutoShapeType shapeType, int handleCount)
        {
            // キャッシュから取得
            if (AngleInterpretationCache.ContainsKey(shapeType))
            {
                var cached = AngleInterpretationCache[shapeType];
                return cached.Take(handleCount).ToList();
            }

            // キャッシュにない場合は作成
            var interpretations = new List<string>();
            switch (shapeType)
            {
                case Office.MsoAutoShapeType.msoShapeArc:
                    interpretations.AddRange(new[] { "開始角度", "終了角度" });
                    break;
                case Office.MsoAutoShapeType.msoShapeChord:
                case Office.MsoAutoShapeType.msoShapePie:
                    interpretations.AddRange(new[] { "開始角度", "終了角度" });
                    break;
                case Office.MsoAutoShapeType.msoShapeBlockArc:
                    interpretations.AddRange(new[] { "開始角度", "終了角度", "内径比率" });
                    break;
                case Office.MsoAutoShapeType.msoShapeDonut:
                    interpretations.Add("内径比率");
                    break;
                case Office.MsoAutoShapeType.msoShapeMoon:
                    interpretations.Add("三日月の角度");
                    break;
                default:
                    for (int i = 0; i < 8; i++) // 最大8個まで
                    {
                        interpretations.Add($"角度{i + 1}");
                    }
                    break;
            }

            // キャッシュに保存
            AngleInterpretationCache[shapeType] = interpretations;

            return interpretations.Take(handleCount).ToList();
        }

        /// <summary>
        /// 図形が角度ハンドルを持つタイプかどうかを判定
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>角度ハンドルを持つ場合true</returns>
        private bool IsAngleHandleShape(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    switch (shape.AutoShapeType)
                    {
                        case Office.MsoAutoShapeType.msoShapeArc:           // 円弧
                        case Office.MsoAutoShapeType.msoShapeChord:         // 弦
                        case Office.MsoAutoShapeType.msoShapePie:           // 扇形（部分円）
                        case Office.MsoAutoShapeType.msoShapeBlockArc:      // ブロック円弧
                        case Office.MsoAutoShapeType.msoShapeDonut:         // ドーナツ
                        case Office.MsoAutoShapeType.msoShapeMoon:          // 三日月
                        case Office.MsoAutoShapeType.msoShapeSmileyFace:    // スマイリーフェイス（角度調整可能）
                            return true;
                        default:
                            return false;
                    }
                },
                $"角度ハンドル判定: {shape.Name}",
                defaultValue: false,
                suppressErrors: true);
        }

        /// <summary>
        /// 図形が調整ハンドルを持つタイプかどうかを判定
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>調整ハンドルを持つ場合true</returns>
        private bool IsAdjustmentHandleShape(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    // 実際の調整ハンドル数をチェック
                    int actualHandleCount = shape.Adjustments.Count;
                    ComExceptionHandler.LogDebug($"IsAdjustmentHandleShape: {shape.Name} - ハンドル数 = {actualHandleCount}");

                    if (actualHandleCount > 0)
                    {
                        ComExceptionHandler.LogDebug($"  → 実際のハンドルが存在するため true");
                        return true;
                    }

                    // 特定の図形タイプをチェック
                    var shapeType = shape.AutoShapeType;
                    ComExceptionHandler.LogDebug($"  → ハンドル数0、タイプチェック: {shapeType}");

                    switch (shapeType)
                    {
                        // 四角形セクション
                        case Office.MsoAutoShapeType.msoShapeRoundedRectangle:
                            ComExceptionHandler.LogDebug($"  → 角丸四角形のため true");
                            return true;
                        case Office.MsoAutoShapeType.msoShapeSnip1Rectangle:
                        case Office.MsoAutoShapeType.msoShapeSnip2SameRectangle:
                        case Office.MsoAutoShapeType.msoShapeSnip2DiagRectangle:
                        case Office.MsoAutoShapeType.msoShapeSnipRoundRectangle:
                        case Office.MsoAutoShapeType.msoShapeRound1Rectangle:
                        case Office.MsoAutoShapeType.msoShapeRound2SameRectangle:
                        case Office.MsoAutoShapeType.msoShapeRound2DiagRectangle:
                            ComExceptionHandler.LogDebug($"  → 四角形セクションのため true");
                            return true;

                        // 基本図形
                        case Office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                        case Office.MsoAutoShapeType.msoShapeParallelogram:
                        case Office.MsoAutoShapeType.msoShapeTrapezoid:
                        case Office.MsoAutoShapeType.msoShapeHexagon:
                        case Office.MsoAutoShapeType.msoShapeOctagon:
                        case Office.MsoAutoShapeType.msoShapeTear:
                        case Office.MsoAutoShapeType.msoShapeFrame:
                        case Office.MsoAutoShapeType.msoShapeHalfFrame:
                        case Office.MsoAutoShapeType.msoShapeCorner:
                        case Office.MsoAutoShapeType.msoShapeDiagonalStripe:
                        case Office.MsoAutoShapeType.msoShapeCross:
                        case Office.MsoAutoShapeType.msoShapePlaque:
                        case Office.MsoAutoShapeType.msoShapeCan:
                        case Office.MsoAutoShapeType.msoShapeCube:
                        case Office.MsoAutoShapeType.msoShapeBevel:
                        case Office.MsoAutoShapeType.msoShapeDonut:
                        case Office.MsoAutoShapeType.msoShapeNoSymbol:
                        case Office.MsoAutoShapeType.msoShapeBlockArc:
                        case Office.MsoAutoShapeType.msoShapeFoldedCorner:
                        case Office.MsoAutoShapeType.msoShapeSmileyFace:
                        case Office.MsoAutoShapeType.msoShapeSun:
                        case Office.MsoAutoShapeType.msoShapeMoon:
                        case Office.MsoAutoShapeType.msoShapeLeftBracket:
                        case Office.MsoAutoShapeType.msoShapeRightBracket:
                        case Office.MsoAutoShapeType.msoShapeLeftBrace:
                        case Office.MsoAutoShapeType.msoShapeRightBrace:

                        // ブロック矢印
                        case Office.MsoAutoShapeType.msoShapeRightArrow:
                        case Office.MsoAutoShapeType.msoShapeLeftArrow:
                        case Office.MsoAutoShapeType.msoShapeUpArrow:
                        case Office.MsoAutoShapeType.msoShapeDownArrow:
                        case Office.MsoAutoShapeType.msoShapeLeftRightArrow:
                        case Office.MsoAutoShapeType.msoShapeUpDownArrow:
                        case Office.MsoAutoShapeType.msoShapeQuadArrow:
                        case Office.MsoAutoShapeType.msoShapeLeftRightUpArrow:
                        case Office.MsoAutoShapeType.msoShapeBentArrow:
                        case Office.MsoAutoShapeType.msoShapeBentUpArrow:
                        case Office.MsoAutoShapeType.msoShapeCurvedRightArrow:
                        case Office.MsoAutoShapeType.msoShapeCurvedLeftArrow:
                        case Office.MsoAutoShapeType.msoShapeCurvedUpArrow:
                        case Office.MsoAutoShapeType.msoShapeCurvedDownArrow:
                        case Office.MsoAutoShapeType.msoShapeStripedRightArrow:
                        case Office.MsoAutoShapeType.msoShapeNotchedRightArrow:
                        case Office.MsoAutoShapeType.msoShapePentagon:
                        case Office.MsoAutoShapeType.msoShapeChevron:

                        // 数式図形
                        case Office.MsoAutoShapeType.msoShapeMathPlus:
                        case Office.MsoAutoShapeType.msoShapeMathMinus:
                        case Office.MsoAutoShapeType.msoShapeMathMultiply:
                        case Office.MsoAutoShapeType.msoShapeMathDivide:
                        case Office.MsoAutoShapeType.msoShapeMathEqual:
                        case Office.MsoAutoShapeType.msoShapeMathNotEqual:

                        // 星とリボン
                        case Office.MsoAutoShapeType.msoShape4pointStar:
                        case Office.MsoAutoShapeType.msoShape5pointStar:
                        case Office.MsoAutoShapeType.msoShape6pointStar:
                        case Office.MsoAutoShapeType.msoShape7pointStar:
                        case Office.MsoAutoShapeType.msoShape8pointStar:
                        case Office.MsoAutoShapeType.msoShape10pointStar:
                        case Office.MsoAutoShapeType.msoShape12pointStar:
                        case Office.MsoAutoShapeType.msoShape16pointStar:
                        case Office.MsoAutoShapeType.msoShape24pointStar:
                        case Office.MsoAutoShapeType.msoShape32pointStar:
                        case Office.MsoAutoShapeType.msoShapeUpRibbon:
                        case Office.MsoAutoShapeType.msoShapeDownRibbon:
                        case Office.MsoAutoShapeType.msoShapeCurvedUpRibbon:
                        case Office.MsoAutoShapeType.msoShapeCurvedDownRibbon:
                        case Office.MsoAutoShapeType.msoShapeVerticalScroll:
                        case Office.MsoAutoShapeType.msoShapeHorizontalScroll:

                        // 吹き出し
                        case Office.MsoAutoShapeType.msoShapeRectangularCallout:
                        case Office.MsoAutoShapeType.msoShapeRoundedRectangularCallout:
                        case Office.MsoAutoShapeType.msoShapeOvalCallout:
                        case Office.MsoAutoShapeType.msoShapeCloudCallout:
                            ComExceptionHandler.LogDebug($"  → 対応図形タイプのため true");
                            return true;

                        default:
                            ComExceptionHandler.LogDebug($"  → 対象外タイプのため false");
                            return false;
                    }
                },
                $"調整ハンドル判定: {shape.Name}",
                defaultValue: false,
                suppressErrors: true);
        }

        /// <summary>
        /// 図形の現在のハンドル情報を取得（詳細版 - ダイアログ表示後に呼び出し）
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>ハンドル情報</returns>
        public ShapeHandleInfo GetHandleInfoDetailed(PowerPoint.Shape shape)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () => {
                    var info = new ShapeHandleInfo
                    {
                        ShapeName = shape.Name,
                        ShapeType = shape.AutoShapeType.ToString(),
                        AdjustmentCount = shape.Adjustments.Count,
                        IsAngleHandleShape = IsAngleHandleShape(shape),
                        IsAdjustmentHandleShape = IsAdjustmentHandleShape(shape),
                        OriginalShape = shape
                    };

                    // 調整ハンドル値を取得（最大8個まで）
                    for (int i = 1; i <= Math.Min(shape.Adjustments.Count, Constants.MAX_SUPPORTED_HANDLES); i++)
                    {
                        float normalizedValue = shape.Adjustments[i];
                        info.AdjustmentValues.Add(normalizedValue);

                        // mm値と度数値も計算して保存
                        if (info.IsAngleHandleShape)
                        {
                            float degreeValue = ConvertNormalizedToDegreeByShapeType(normalizedValue, shape.AutoShapeType, i - 1);
                            info.AdjustmentValuesInDegree.Add(degreeValue);
                        }
                        else
                        {
                            info.AdjustmentValuesInMm.Add(ConvertNormalizedToMm(normalizedValue, shape, i - 1));
                        }
                    }

                    // 角度ハンドル図形の場合、角度の意味を解釈
                    if (info.IsAngleHandleShape && info.AdjustmentValues.Count > 0)
                    {
                        info.AngleInterpretation = GetAngleInterpretation(shape.AutoShapeType, info.AdjustmentValues.Count);
                    }

                    // 調整ハンドル図形の場合、調整の意味を解釈
                    if (info.IsAdjustmentHandleShape && info.AdjustmentValues.Count > 0)
                    {
                        info.AdjustmentInterpretation = GetAdjustmentHandleInterpretation(shape.AutoShapeType, info.AdjustmentValues.Count);
                    }

                    return info;
                },
                $"ハンドル情報取得（詳細）: {shape.Name}",
                defaultValue: new ShapeHandleInfo { ShapeName = "エラー" },
                suppressErrors: true);
        }

        /// <summary>
        /// 図形タイプに応じた角度ハンドルの意味を取得
        /// </summary>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="handleCount">ハンドル数</param>
        /// <returns>角度ハンドルの意味</returns>
        private List<string> GetAngleInterpretation(Office.MsoAutoShapeType shapeType, int handleCount)
        {
            var interpretations = new List<string>();

            switch (shapeType)
            {
                case Office.MsoAutoShapeType.msoShapeArc:
                    if (handleCount >= 1) interpretations.Add("開始角度");
                    if (handleCount >= 2) interpretations.Add("終了角度");
                    break;

                case Office.MsoAutoShapeType.msoShapeChord:
                case Office.MsoAutoShapeType.msoShapePie:
                    if (handleCount >= 1) interpretations.Add("開始角度");
                    if (handleCount >= 2) interpretations.Add("終了角度");
                    break;

                case Office.MsoAutoShapeType.msoShapeBlockArc:
                    if (handleCount >= 1) interpretations.Add("開始角度");
                    if (handleCount >= 2) interpretations.Add("終了角度");
                    if (handleCount >= 3) interpretations.Add("内径比率");
                    break;

                case Office.MsoAutoShapeType.msoShapeDonut:
                    if (handleCount >= 1) interpretations.Add("内径比率");
                    break;

                case Office.MsoAutoShapeType.msoShapeMoon:
                    if (handleCount >= 1) interpretations.Add("三日月の角度");
                    break;

                default:
                    for (int i = 0; i < handleCount; i++)
                    {
                        interpretations.Add($"角度{i + 1}");
                    }
                    break;
            }

            return interpretations;
        }

        /// <summary>
        /// 図形タイプに応じた調整ハンドルの意味を取得
        /// </summary>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="handleCount">ハンドル数</param>
        /// <returns>調整ハンドルの意味</returns>
        private List<string> GetAdjustmentHandleInterpretation(Office.MsoAutoShapeType shapeType, int handleCount)
        {
            var interpretations = new List<string>();

            switch (shapeType)
            {
                // 基本図形
                case Office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                    if (handleCount >= 1) interpretations.Add("三角形の高さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeParallelogram:
                    if (handleCount >= 1) interpretations.Add("平行四辺形の傾き");
                    break;

                case Office.MsoAutoShapeType.msoShapeTrapezoid:
                    if (handleCount >= 1) interpretations.Add("台形の傾き");
                    break;

                case Office.MsoAutoShapeType.msoShapeHexagon:
                case Office.MsoAutoShapeType.msoShapeOctagon:
                    if (handleCount >= 1) interpretations.Add("多角形の形状");
                    break;

                case Office.MsoAutoShapeType.msoShapeTear:
                    if (handleCount >= 1) interpretations.Add("涙形の曲がり");
                    break;

                case Office.MsoAutoShapeType.msoShapeFrame:
                case Office.MsoAutoShapeType.msoShapeHalfFrame:
                    if (handleCount >= 1) interpretations.Add("フレームの太さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeCorner:
                    if (handleCount >= 1) interpretations.Add("L字の角度");
                    break;

                case Office.MsoAutoShapeType.msoShapeDiagonalStripe:
                    if (handleCount >= 1) interpretations.Add("縞の角度");
                    break;

                case Office.MsoAutoShapeType.msoShapeCross:
                    if (handleCount >= 1) interpretations.Add("十字の太さ");
                    break;

                case Office.MsoAutoShapeType.msoShapePlaque:
                    if (handleCount >= 1) interpretations.Add("プラークの深さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeCan:
                    if (handleCount >= 1) interpretations.Add("円柱の高さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeCube:
                    if (handleCount >= 1) interpretations.Add("直方体の奥行き");
                    break;

                case Office.MsoAutoShapeType.msoShapeBevel:
                    if (handleCount >= 1) interpretations.Add("ベベルの深さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeDonut:
                    if (handleCount >= 1) interpretations.Add("内径の比率");
                    break;

                case Office.MsoAutoShapeType.msoShapeNoSymbol:
                    if (handleCount >= 1) interpretations.Add("禁止線の太さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeBlockArc:
                    if (handleCount >= 1) interpretations.Add("開始角度");
                    if (handleCount >= 2) interpretations.Add("終了角度");
                    if (handleCount >= 3) interpretations.Add("内径比率");
                    break;

                case Office.MsoAutoShapeType.msoShapeFoldedCorner:
                    if (handleCount >= 1) interpretations.Add("折り角のサイズ");
                    break;

                case Office.MsoAutoShapeType.msoShapeSmileyFace:
                    if (handleCount >= 1) interpretations.Add("笑顔の度合い");
                    break;

                case Office.MsoAutoShapeType.msoShapeSun:
                    if (handleCount >= 1) interpretations.Add("光線の長さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeMoon:
                    if (handleCount >= 1) interpretations.Add("三日月の角度");
                    break;

                // 四角形セクション（新規追加）
                case Office.MsoAutoShapeType.msoShapeRoundedRectangle:
                    if (handleCount >= 1) interpretations.Add("角丸の半径");
                    break;

                case Office.MsoAutoShapeType.msoShapeSnip1Rectangle:
                    if (handleCount >= 1) interpretations.Add("切り取り角のサイズ");
                    break;

                case Office.MsoAutoShapeType.msoShapeSnip2SameRectangle:
                    if (handleCount >= 1) interpretations.Add("切り取り角のサイズ");
                    break;

                case Office.MsoAutoShapeType.msoShapeSnip2DiagRectangle:
                    if (handleCount >= 1) interpretations.Add("切り取り角のサイズ");
                    break;

                case Office.MsoAutoShapeType.msoShapeSnipRoundRectangle:
                    if (handleCount >= 1) interpretations.Add("切り取り角のサイズ");
                    if (handleCount >= 2) interpretations.Add("角丸の半径");
                    break;

                case Office.MsoAutoShapeType.msoShapeRound1Rectangle:
                    if (handleCount >= 1) interpretations.Add("角丸の半径");
                    break;

                case Office.MsoAutoShapeType.msoShapeRound2SameRectangle:
                    if (handleCount >= 1) interpretations.Add("角丸の半径");
                    break;

                case Office.MsoAutoShapeType.msoShapeRound2DiagRectangle:
                    if (handleCount >= 1) interpretations.Add("角丸の半径");
                    break;

                // 括弧
                case Office.MsoAutoShapeType.msoShapeLeftBracket:
                case Office.MsoAutoShapeType.msoShapeRightBracket:
                    if (handleCount >= 1) interpretations.Add("大かっこの曲がり");
                    break;

                case Office.MsoAutoShapeType.msoShapeLeftBrace:
                case Office.MsoAutoShapeType.msoShapeRightBrace:
                    if (handleCount >= 1) interpretations.Add("中かっこの曲がり");
                    break;

                // ブロック矢印
                case Office.MsoAutoShapeType.msoShapeRightArrow:
                case Office.MsoAutoShapeType.msoShapeLeftArrow:
                case Office.MsoAutoShapeType.msoShapeUpArrow:
                case Office.MsoAutoShapeType.msoShapeDownArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeLeftRightArrow:
                case Office.MsoAutoShapeType.msoShapeUpDownArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeQuadArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeLeftRightUpArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeBentArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    if (handleCount >= 3) interpretations.Add("曲がり角の位置");
                    break;

                case Office.MsoAutoShapeType.msoShapeBentUpArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    if (handleCount >= 3) interpretations.Add("上向き折線の幅");
                    break;

                case Office.MsoAutoShapeType.msoShapeCurvedRightArrow:
                case Office.MsoAutoShapeType.msoShapeCurvedLeftArrow:
                case Office.MsoAutoShapeType.msoShapeCurvedUpArrow:
                case Office.MsoAutoShapeType.msoShapeCurvedDownArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    if (handleCount >= 3) interpretations.Add("カーブの度合い");
                    break;

                case Office.MsoAutoShapeType.msoShapeStripedRightArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    if (handleCount >= 3) interpretations.Add("ストライプの位置");
                    break;

                case Office.MsoAutoShapeType.msoShapeNotchedRightArrow:
                    if (handleCount >= 1) interpretations.Add("矢じりの幅");
                    if (handleCount >= 2) interpretations.Add("矢じりの長さ");
                    if (handleCount >= 3) interpretations.Add("V字の深さ");
                    break;

                case Office.MsoAutoShapeType.msoShapePentagon:
                    if (handleCount >= 1) interpretations.Add("五角形の形状");
                    break;

                case Office.MsoAutoShapeType.msoShapeChevron:
                    if (handleCount >= 1) interpretations.Add("山形の角度");
                    break;

                // 数式図形
                case Office.MsoAutoShapeType.msoShapeMathPlus:
                    if (handleCount >= 1) interpretations.Add("十字の太さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeMathMinus:
                    if (handleCount >= 1) interpretations.Add("横線の太さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeMathMultiply:
                    if (handleCount >= 1) interpretations.Add("×印の太さ");
                    break;

                case Office.MsoAutoShapeType.msoShapeMathDivide:
                    if (handleCount >= 1) interpretations.Add("÷印の間隔");
                    break;

                case Office.MsoAutoShapeType.msoShapeMathEqual:
                    if (handleCount >= 1) interpretations.Add("等号の間隔");
                    break;

                case Office.MsoAutoShapeType.msoShapeMathNotEqual:
                    if (handleCount >= 1) interpretations.Add("不等号の角度");
                    break;

                // 星とリボン
                case Office.MsoAutoShapeType.msoShape4pointStar:
                case Office.MsoAutoShapeType.msoShape5pointStar:
                case Office.MsoAutoShapeType.msoShape6pointStar:
                case Office.MsoAutoShapeType.msoShape7pointStar:
                case Office.MsoAutoShapeType.msoShape8pointStar:
                case Office.MsoAutoShapeType.msoShape10pointStar:
                case Office.MsoAutoShapeType.msoShape12pointStar:
                case Office.MsoAutoShapeType.msoShape16pointStar:
                case Office.MsoAutoShapeType.msoShape24pointStar:
                case Office.MsoAutoShapeType.msoShape32pointStar:
                    if (handleCount >= 1) interpretations.Add("内側の頂点の位置");
                    break;

                case Office.MsoAutoShapeType.msoShapeUpRibbon:
                case Office.MsoAutoShapeType.msoShapeDownRibbon:
                    if (handleCount >= 1) interpretations.Add("リボンの幅");
                    if (handleCount >= 2) interpretations.Add("リボンの形状");
                    break;

                case Office.MsoAutoShapeType.msoShapeCurvedUpRibbon:
                case Office.MsoAutoShapeType.msoShapeCurvedDownRibbon:
                    if (handleCount >= 1) interpretations.Add("リボンの幅");
                    if (handleCount >= 2) interpretations.Add("リボンのカーブ");
                    break;

                case Office.MsoAutoShapeType.msoShapeVerticalScroll:
                case Office.MsoAutoShapeType.msoShapeHorizontalScroll:
                    if (handleCount >= 1) interpretations.Add("スクロールの巻き");
                    break;

                // 吹き出し
                case Office.MsoAutoShapeType.msoShapeRectangularCallout:
                case Office.MsoAutoShapeType.msoShapeRoundedRectangularCallout:
                case Office.MsoAutoShapeType.msoShapeOvalCallout:
                case Office.MsoAutoShapeType.msoShapeCloudCallout:
                    if (handleCount >= 1) interpretations.Add("吹き出し線の横位置");
                    if (handleCount >= 2) interpretations.Add("吹き出し線の縦位置");
                    break;

                default:
                    for (int i = 0; i < handleCount; i++)
                    {
                        interpretations.Add($"調整値{i + 1}");
                    }
                    break;
            }

            return interpretations;
        }

        #endregion

        #region デバッグ機能

        /// <summary>
        /// 図形の詳細情報をデバッグ出力する
        /// </summary>
        /// <param name="shape">対象図形</param>
        public void DebugShapeInfo(PowerPoint.Shape shape)
        {
            ComExceptionHandler.ExecuteComOperation(
                () => {
                    ComExceptionHandler.LogDebug($"=== 図形詳細情報 ===");
                    ComExceptionHandler.LogDebug($"名前: {shape.Name}");
                    ComExceptionHandler.LogDebug($"タイプ: {shape.Type}");
                    ComExceptionHandler.LogDebug($"オートシェイプタイプ: {shape.AutoShapeType}");
                    ComExceptionHandler.LogDebug($"調整ハンドル数: {shape.Adjustments.Count}");

                    // 調整ハンドル値を出力
                    for (int i = 1; i <= shape.Adjustments.Count; i++)
                    {
                        ComExceptionHandler.LogDebug($"ハンドル{i}: {shape.Adjustments[i]}");
                    }

                    ComExceptionHandler.LogDebug($"IsAdjustmentHandleShape判定: {IsAdjustmentHandleShape(shape)}");
                    ComExceptionHandler.LogDebug($"IsAngleHandleShape判定: {IsAngleHandleShape(shape)}");
                    ComExceptionHandler.LogDebug($"==================");
                },
                "図形詳細情報出力",
                suppressErrors: true);
        }

        /// <summary>
        /// 複数図形の詳細情報をデバッグ出力する（軽量版）
        /// </summary>
        /// <param name="shapes">対象図形リスト</param>
        public void DebugMultipleShapesInfoLight(List<PowerPoint.Shape> shapes)
        {
            ComExceptionHandler.LogInfo($"=== 複数図形分析開始 ({shapes?.Count ?? 0}個) ===");

            if (shapes != null)
            {
                for (int i = 0; i < shapes.Count; i++)
                {
                    ComExceptionHandler.LogDebug($"--- 図形 {i + 1} ---");
                    // 軽量版の情報取得のみ
                    var info = GetHandleInfoFast(shapes[i]);
                    ComExceptionHandler.LogDebug($"名前: {info.ShapeName}, タイプ: {info.ShapeType}, ハンドル数: {info.AdjustmentCount}");
                }
            }

            ComExceptionHandler.LogDebug($"=== 複数図形分析終了 ===");
        }

        #endregion
    }

    /// <summary>
    /// 図形ハンドル情報を格納するクラス（拡張版）
    /// </summary>
    public class ShapeHandleInfo
    {
        public string ShapeName { get; set; }
        public string ShapeType { get; set; }
        public int AdjustmentCount { get; set; }
        public bool IsAngleHandleShape { get; set; }
        public bool IsAdjustmentHandleShape { get; set; }
        public List<float> AdjustmentValues { get; set; } = new List<float>();
        public List<float> AdjustmentValuesInMm { get; set; } = new List<float>();
        public List<float> AdjustmentValuesInDegree { get; set; } = new List<float>();
        public List<string> AngleInterpretation { get; set; } = new List<string>();
        public List<string> AdjustmentInterpretation { get; set; } = new List<string>();
        public PowerPoint.Shape OriginalShape { get; set; }

        /// <summary>
        /// 図形タイプの表示名を取得（拡張版）
        /// </summary>
        public string GetDisplayShapeType()
        {
            switch (ShapeType)
            {
                // 基本図形
                case "msoShapeIsoscelesTriangle": return "二等辺三角形";
                case "msoShapeParallelogram": return "平行四辺形";
                case "msoShapeTrapezoid": return "台形";
                case "msoShapeHexagon": return "六角形";
                case "msoShapeOctagon": return "八角形";
                case "msoShapeTear": return "涙形";
                case "msoShapeFrame": return "フレーム";
                case "msoShapeHalfFrame": return "フレーム（半分）";
                case "msoShapeCorner": return "L字";
                case "msoShapeDiagonalStripe": return "斜め縞";
                case "msoShapeCross": return "十字形";
                case "msoShapePlaque": return "プラーク";
                case "msoShapeCan": return "円柱";
                case "msoShapeCube": return "直方体";
                case "msoShapeBevel": return "ベベル";
                case "msoShapeDonut": return "円：塗りつぶしなし";
                case "msoShapeNoSymbol": return "禁止マーク";
                case "msoShapeBlockArc": return "アーチ";
                case "msoShapeFoldedCorner": return "四角形：角度付き";
                case "msoShapeSmileyFace": return "スマイル";
                case "msoShapeSun": return "太陽";
                case "msoShapeMoon": return "月";
                case "msoShapeLeftBracket": return "左大かっこ";
                case "msoShapeRightBracket": return "右大かっこ";
                case "msoShapeLeftBrace": return "左中かっこ";
                case "msoShapeRightBrace": return "右中かっこ";

                // 四角形セクション（新規追加）
                case "msoShapeRoundedRectangle": return "四角形：角を丸くする";
                case "msoShapeSnip1Rectangle": return "四角形：１つの角を切り取る";
                case "msoShapeSnip2SameRectangle": return "四角形：2つの角を切り取る";
                case "msoShapeSnip2DiagRectangle": return "四角形：対角を切り取る";
                case "msoShapeSnipRoundRectangle": return "四角形：１つの角を切り取り１つの角を丸める";
                case "msoShapeRound1Rectangle": return "四角形：１つの角を丸める";
                case "msoShapeRound2SameRectangle": return "四角形：上の2つの角を丸める";
                case "msoShapeRound2DiagRectangle": return "四角形：対角を丸める";

                // ブロック矢印
                case "msoShapeRightArrow": return "矢印：右";
                case "msoShapeLeftArrow": return "矢印：左";
                case "msoShapeUpArrow": return "矢印：上";
                case "msoShapeDownArrow": return "矢印：下";
                case "msoShapeLeftRightArrow": return "矢印：左右";
                case "msoShapeUpDownArrow": return "矢印：上下";
                case "msoShapeQuadArrow": return "矢印：四方向";
                case "msoShapeLeftRightUpArrow": return "矢印：三方向";
                case "msoShapeBentArrow": return "矢印：二方向";
                case "msoShapeBentUpArrow": return "矢印：上向き折線";
                case "msoShapeCurvedRightArrow": return "矢印：右カーブ";
                case "msoShapeCurvedLeftArrow": return "矢印：左カーブ";
                case "msoShapeCurvedUpArrow": return "矢印：上カーブ";
                case "msoShapeCurvedDownArrow": return "矢印：下カーブ";
                case "msoShapeStripedRightArrow": return "矢印：ストライプ";
                case "msoShapeNotchedRightArrow": return "矢印：V字型";
                case "msoShapePentagon": return "矢印：五方向";
                case "msoShapeChevron": return "矢印：山形";

                // 数式図形
                case "msoShapeMathPlus": return "加算記号";
                case "msoShapeMathMinus": return "減算記号";
                case "msoShapeMathMultiply": return "乗算記号";
                case "msoShapeMathDivide": return "除算記号";
                case "msoShapeMathEqual": return "次の値と等しい";
                case "msoShapeMathNotEqual": return "等号否定";

                // 星とリボン
                case "msoShape4pointStar": return "星：4pt";
                case "msoShape5pointStar": return "星：5pt";
                case "msoShape6pointStar": return "星：6pt";
                case "msoShape7pointStar": return "星：7pt";
                case "msoShape8pointStar": return "星：8pt";
                case "msoShape10pointStar": return "星：10pt";
                case "msoShape12pointStar": return "星：12pt";
                case "msoShape16pointStar": return "星：16pt";
                case "msoShape24pointStar": return "星：24pt";
                case "msoShape32pointStar": return "星：32pt";
                case "msoShapeUpRibbon": return "リボン：上に曲がる";
                case "msoShapeDownRibbon": return "リボン：下に曲がる";
                case "msoShapeCurvedUpRibbon": return "リボン：カーブして上方向に曲がる";
                case "msoShapeCurvedDownRibbon": return "リボン：カーブして下方向に曲がる";
                case "msoShapeVerticalScroll": return "スクロール：縦";
                case "msoShapeHorizontalScroll": return "スクロール：横";

                // 吹き出し
                case "msoShapeRectangularCallout": return "吹き出し：四角形";
                case "msoShapeRoundedRectangularCallout": return "吹き出し：角を丸めた四角形";
                case "msoShapeOvalCallout": return "吹き出し：円形";
                case "msoShapeCloudCallout": return "思考の吹き出し：雲形";

                // 角度ハンドル図形
                case "msoShapeArc": return "円弧";
                case "msoShapeChord": return "弦";
                case "msoShapePie": return "扇形（部分円）";

                default: return ShapeType;
            }
        }

        public override string ToString()
        {
            string typeInfo = IsAngleHandleShape ? $"{GetDisplayShapeType()}（角度ハンドル）" :
                             IsAdjustmentHandleShape ? $"{GetDisplayShapeType()}（調整ハンドル）" :
                             GetDisplayShapeType();
            return $"{ShapeName} ({typeInfo}): ハンドル数 {AdjustmentCount}";
        }
    }

    /// <summary>
    /// 図形群の分析結果を格納するクラス
    /// </summary>
    public class ShapeHandleAnalysis
    {
        public int TotalShapes { get; set; }
        public int ShapesWithAdjustmentHandles { get; set; }
        public int ShapesWithAngleHandles { get; set; }
        public int MaxAdjustmentHandles { get; set; }
        public int RecommendedHandleCount { get; set; }
        public int RecommendedAngleHandleCount { get; set; }
        public List<ShapeHandleInfo> ShapeInfos { get; set; } = new List<ShapeHandleInfo>();
        public List<string> AngleHandleShapeTypes { get; set; } = new List<string>();

        /// <summary>
        /// 調整ハンドルを持つ図形があるか
        /// </summary>
        public bool HasAdjustmentHandles => ShapesWithAdjustmentHandles > 0;

        /// <summary>
        /// 角度ハンドルを持つ図形があるか
        /// </summary>
        public bool HasAngleHandles => ShapesWithAngleHandles > 0;

        /// <summary>
        /// 統一されたハンドル数を持つか
        /// </summary>
        public bool HasUniformHandleCount =>
            ShapeInfos.Where(info => info.AdjustmentCount > 0)
                     .Select(info => info.AdjustmentCount)
                     .Distinct()
                     .Count() <= 1;

        /// <summary>
        /// 角度ハンドル図形のみが選択されているか
        /// </summary>
        public bool IsAngleHandleOnly =>
            HasAngleHandles && ShapesWithAngleHandles == ShapesWithAdjustmentHandles;

        public override string ToString()
        {
            return $"総数: {TotalShapes}, 調整ハンドル有り: {ShapesWithAdjustmentHandles}, " +
                   $"角度ハンドル有り: {ShapesWithAngleHandles}, 推奨ハンドル数: {RecommendedHandleCount}";
        }
    }
}