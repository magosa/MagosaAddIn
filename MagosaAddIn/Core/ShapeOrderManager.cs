using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形の処理順序を管理するクラス。
    /// ユーザーが指定した順序でレイヤー調整またはスタック保存を実行する。
    /// </summary>
    public class ShapeOrderManager
    {
        /// <summary>
        /// 指定した順序で図形を選び直す。
        /// Z 順（重なり順）は変更せず、Selection.ShapeRange のクリック順のみを更新する。
        /// これにより後続のナンバリング等の操作が指定順で実行される。
        /// </summary>
        /// <param name="orderedShapes">新たな選択順に並んだ図形リスト</param>
        public void ReSelectInOrder(List<PowerPoint.Shape> orderedShapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (orderedShapes == null || orderedShapes.Count == 0)
                        throw new ArgumentException("再選択する図形が指定されていません。");

                    // 最初の図形を単独選択（現在の選択を解除）
                    orderedShapes[0].Select(Office.MsoTriState.msoTrue);

                    // 2番目以降を追加選択（クリック順として記録される）
                    for (int i = 1; i < orderedShapes.Count; i++)
                    {
                        orderedShapes[i].Select(Office.MsoTriState.msoFalse);
                    }

                    ComExceptionHandler.LogDebug($"選択順序で再選択完了: {orderedShapes.Count}個");
                },
                "選択順序で再選択");
        }

        /// <summary>
        /// 指定した順序で図形のレイヤー（Z順）を適用する。
        /// orderedShapes[0] が最背面、orderedShapes[Last] が最前面となる。
        /// ※ 体裁に影響するため、通常は ReSelectInOrder を使用すること。
        /// </summary>
        /// <param name="orderedShapes">適用順に並んだ図形リスト（先頭=背面）</param>
        public void ApplyLayerOrder(List<PowerPoint.Shape> orderedShapes)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (orderedShapes == null || orderedShapes.Count < 2)
                        throw new ArgumentException("レイヤー適用には2個以上の図形が必要です。");

                    // 最初の図形を最背面に配置
                    orderedShapes[0].ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                    // 2番目以降を順番に前面へ配置
                    for (int i = 1; i < orderedShapes.Count; i++)
                    {
                        orderedShapes[i].ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    }

                    ComExceptionHandler.LogDebug($"選択順序レイヤー適用完了: {orderedShapes.Count}個");
                },
                "選択順序レイヤー適用");
        }

        /// <summary>
        /// 指定した順序で図形をスタックに保存する。
        /// </summary>
        /// <param name="orderedShapes">保存する図形リスト</param>
        /// <param name="stack">保存先のShapeStackインスタンス</param>
        /// <returns>作成されたスタックID</returns>
        public int SaveToStack(List<PowerPoint.Shape> orderedShapes, ShapeStack stack)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    if (orderedShapes == null || orderedShapes.Count == 0)
                        throw new ArgumentException("スタックに保存する図形が指定されていません。");
                    if (stack == null)
                        throw new ArgumentNullException(nameof(stack));

                    int stackId = stack.PushStack(orderedShapes);
                    ComExceptionHandler.LogDebug($"選択順序スタック保存完了: Stack {stackId}, 図形数: {orderedShapes.Count}");
                    return stackId;
                },
                "選択順序スタック保存");
        }

        /// <summary>
        /// 図形の種類を日本語で返す。
        /// </summary>
        public static string GetShapeTypeName(PowerPoint.Shape shape)
        {
            try
            {
                switch (shape.Type)
                {
                    case Office.MsoShapeType.msoAutoShape:
                        return GetAutoShapeTypeName(shape);
                    case Office.MsoShapeType.msoPicture:
                        return "画像";
                    case Office.MsoShapeType.msoLinkedPicture:
                        return "リンク画像";
                    case Office.MsoShapeType.msoTextBox:
                        return "テキストボックス";
                    case Office.MsoShapeType.msoGroup:
                        return "グループ";
                    case Office.MsoShapeType.msoLine:
                        return "線";
                    case Office.MsoShapeType.msoFreeform:
                        return "フリーフォーム";
                    case Office.MsoShapeType.msoTable:
                        return "表";
                    case Office.MsoShapeType.msoChart:
                        return "グラフ";
                    default:
                        return "図形";
                }
            }
            catch
            {
                return "図形";
            }
        }

        private static string GetAutoShapeTypeName(PowerPoint.Shape shape)
        {
            try
            {
                switch (shape.AutoShapeType)
                {
                    case Office.MsoAutoShapeType.msoShapeRectangle:
                        return "四角形";
                    case Office.MsoAutoShapeType.msoShapeRoundedRectangle:
                        return "角丸四角形";
                    case Office.MsoAutoShapeType.msoShapeOval:
                        return "楕円";
                    case Office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                    case Office.MsoAutoShapeType.msoShapeRightTriangle:
                        return "三角形";
                    case Office.MsoAutoShapeType.msoShapeParallelogram:
                        return "平行四辺形";
                    case Office.MsoAutoShapeType.msoShapeHexagon:
                        return "六角形";
                    case Office.MsoAutoShapeType.msoShapeDiamond:
                        return "ひし形";
                    default:
                        return "オートシェイプ";
                }
            }
            catch
            {
                return "オートシェイプ";
            }
        }
    }
}
