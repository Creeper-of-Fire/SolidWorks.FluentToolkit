using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Utils;
using System.Collections.Generic;
using System.Linq;

namespace SolidWorks.Helpers;

/// <summary>
/// 为 ISketchManager 提供高级绘图功能的扩展方法。
/// 这些方法旨在通过接受全局3D坐标来简化草图绘制，并将坐标转换的复杂性完全封装。
/// </summary>
public static class SketchHelper
{
    private static SldWorks SldWorks { get; } = SolidWorksConnector.GetSldWorksApp();

    /// <summary>
    /// 根据一系列全局3D坐标点，在当前激活的草图中绘制一个轮廓。
    /// 这个方法会自动处理从3D模型空间到2D草图空间的坐标转换。
    /// </summary>
    /// <param name="sm">要进行操作的 SketchManager。</param>
    /// <param name="points">定义轮廓顶点的一系列全局3D坐标点。</param>
    /// <param name="close">如果为 true，则在最后一个点和第一个点之间创建一条线，以闭合轮廓。</param>
    /// <returns>返回传入的 SketchManager 实例，以支持链式调用。</returns>
    /// <exception cref="ArgumentException">当提供的点少于2个时抛出。</exception>
    /// <exception cref="InvalidOperationException">当无法获取有效的草图变换矩阵时抛出。</exception>
    public static ISketchManager DrawContour(
        this ISketchManager sm,
        IEnumerable<Point3D> points,
        bool close = true)
    {
        // 步骤 1: 数据准备和验证
        var pointList = points.ToList();
        if (pointList.Count < 2)
        {
            throw new ArgumentException("绘制轮廓至少需要2个点。", nameof(points));
        }

        // 步骤 2: 获取从模型空间到草图空间的变换矩阵 (这是核心)
        var activeSketch = sm.ActiveSketch.AssertNotNull("当前没有激活的草图。");
        var sketchTransform = activeSketch.ModelToSketchTransform.AssertNotNull("无法获取草图的变换矩阵。");
        var mathUtil = SldWorks.IGetMathUtility();

        // 辅助函数：将全局3D点转换为草图2D点
        double[] ToSketchCoords(Point3D p)
        {
            var pointData = new[] { p.X, p.Y, p.Z };
            var comObject = mathUtil.CreatePoint(pointData);
            IMathPoint mathPoint = (comObject as IMathPoint)
                .AssertNotNull("从坐标创建 IMathPoint 失败。");
            IMathPoint transformedPoint = (mathPoint.MultiplyTransform(sketchTransform) as IMathPoint)
                .AssertNotNull("从变换矩阵创建 IMathPoint 失败。");
            return (double[])transformedPoint.ArrayData;
        }

        // 步骤 3: 依次连接所有点
        for (int i = 0; i < pointList.Count - 1; i++)
        {
            var start = ToSketchCoords(pointList[i]);
            var end = ToSketchCoords(pointList[i + 1]);
            sm.CreateLine(start[0], start[1], start[2], end[0], end[1], end[2]);
        }

        // 步骤 4: 如果需要，闭合轮廓
        if (close)
        {
            var last = ToSketchCoords(pointList[^1]); // C# 8.0 索引器语法，获取最后一个元素
            var first = ToSketchCoords(pointList[0]);
            sm.CreateLine(last[0], last[1], last[2], first[0], first[1], first[2]);
        }

        return sm;
    }
}