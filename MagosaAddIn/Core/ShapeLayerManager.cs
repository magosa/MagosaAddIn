using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MagosaAddIn.Core
{
    /// <summary>
    /// 図形のレイヤー（重なり順）を調整するクラス
    /// </summary>
    public class ShapeLayerManager
    {
        /// <summary>
        /// 図形のレイヤーを調整
        /// </summary>
        /// <param name="shapes">調整する図形のリスト（選択順）</param>
        /// <param name="order">調整方向</param>
        public void AdjustLayers(List<PowerPoint.Shape> shapes, LayerOrder order)
        {
            ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    ErrorHandler.ValidateShapes(shapes, Constants.MIN_SHAPES_FOR_LAYER, "レイヤー調整");

                    switch (order)
                    {
                        case LayerOrder.SelectionOrderToFront:
                            AdjustLayersBySelectionOrderToFront(shapes);
                            break;
                        case LayerOrder.SelectionOrderToBack:
                            AdjustLayersBySelectionOrderToBack(shapes);
                            break;
                        case LayerOrder.LeftToRightToFront:
                            AdjustLayersByPositionToFront(shapes, true);
                            break;
                        case LayerOrder.TopToBottomToFront:
                            AdjustLayersByPositionToFront(shapes, false);
                            break;
                        default:
                            throw new ArgumentException($"未対応のレイヤー調整方向: {order}");
                    }
                },
                "レイヤー調整");
        }

        /// <summary>
        /// 選択順に前面へ配置
        /// </summary>
        private void AdjustLayersBySelectionOrderToFront(List<PowerPoint.Shape> shapes)
        {
            // 最初の図形を最背面に配置
            shapes[0].ZOrder(Office.MsoZOrderCmd.msoSendToBack);

            // 2番目以降の図形を順番に前面に配置
            for (int i = 1; i < shapes.Count; i++)
            {
                shapes[i].ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }

            ComExceptionHandler.LogDebug($"選択順に前面配置完了: {shapes.Count}個");
        }

        /// <summary>
        /// 選択順に背面へ配置
        /// </summary>
        private void AdjustLayersBySelectionOrderToBack(List<PowerPoint.Shape> shapes)
        {
            // 最初の図形を最前面に配置
            shapes[0].ZOrder(Office.MsoZOrderCmd.msoBringToFront);

            // 2番目以降の図形を順番に背面に配置
            for (int i = 1; i < shapes.Count; i++)
            {
                shapes[i].ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }

            ComExceptionHandler.LogDebug($"選択順に背面配置完了: {shapes.Count}個");
        }

        /// <summary>
        /// 位置に基づいて前面へ配置
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="isHorizontal">true=左から右、false=上から下</param>
        private void AdjustLayersByPositionToFront(List<PowerPoint.Shape> shapes, bool isHorizontal)
        {
            // 位置でソート
            var sortedShapes = isHorizontal
                ? shapes.OrderBy(s => s.Left).ToList()
                : shapes.OrderBy(s => s.Top).ToList();

            // 最初の図形（最も左/上）を最背面に配置
            sortedShapes[0].ZOrder(Office.MsoZOrderCmd.msoSendToBack);

            // 残りの図形を順番に前面に配置
            for (int i = 1; i < sortedShapes.Count; i++)
            {
                sortedShapes[i].ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }

            string direction = isHorizontal ? "左から右" : "上から下";
            ComExceptionHandler.LogDebug($"{direction}へ前面配置完了: {shapes.Count}個");
        }

        /// <summary>
        /// 現在の図形の重なり順序を取得（デバッグ用）
        /// </summary>
        public List<(string Name, int ZOrder)> GetCurrentZOrders(List<PowerPoint.Shape> shapes)
        {
            return ComExceptionHandler.ExecuteComOperation(
                () =>
                {
                    var result = new List<(string Name, int ZOrder)>();
                    foreach (var shape in shapes)
                    {
                        result.Add((shape.Name, shape.ZOrderPosition));
                    }
                    return result.OrderBy(x => x.ZOrder).ToList();
                },
                "重なり順序取得",
                defaultValue: new List<(string, int)>(),
                suppressErrors: true);
        }
    }
}