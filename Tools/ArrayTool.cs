using SolidWorks.Helpers;
using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;

namespace SolidWorks.Tools;

public static class ArrayTool
{
    /// <summary>
    /// 创建一组汇聚于单个焦点的扇形扫描切除孔。
    /// 这是一个高级辅助函数，它将“在哪里打孔”和“如何打孔”的逻辑解耦。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="entryPoints">一个包含所有孔入口3D坐标的集合。</param>
    /// <param name="focusPoint">所有孔共同指向的3D焦点坐标。</param>
    /// <param name="holeDiameter">每个孔的直径。</param>
    /// <param name="createdPathSketches">（输出）所有为路径创建的临时3D草图特征列表。</param>
    /// <returns>返回传入的 IModelDoc2 对象，以支持链式调用。</returns>
    public static IModelDoc2 CreateConvergingSweptCuts(
        this IModelDoc2 model,
        IEnumerable<Point3D> entryPoints,
        Point3D focusPoint,
        double holeDiameter,
        out List<Feature> createdPathSketches)
    {
        createdPathSketches = [];
        int holeCounter = 0;
        var entryPointList = entryPoints.ToList();
        int totalHoles = entryPointList.Count;

        foreach (var startPoint in entryPointList)
        {
            holeCounter++;
            Console.WriteLine($"日志：(通用阵列) 正在创建第 {holeCounter} / {totalHoles} 个准直孔...");

            // 调用我们已经验证过的3D草图创建函数
            model.Create3DSketchFeature(sm =>
            {
                sm.CreateLine(
                    startPoint.X, startPoint.Y, startPoint.Z,
                    focusPoint.X, focusPoint.Y, focusPoint.Z
                );
            }, out var pathFeature);

            createdPathSketches.Add(pathFeature);

            // 调用扫描切除函数
            var pathSketch = (ISketch)pathFeature.GetSpecificFeature2();
            model.CreateSweptCutWithCircularProfile(pathSketch, holeDiameter, out var cutFeature);

            Console.WriteLine($"  -> 从 ({startPoint.X:F2}, {startPoint.Y:F2}, {startPoint.Z:F2}) -> 焦点。");
            Console.WriteLine($"  -> 特征 '{cutFeature.Name}' 已创建。");
        }

        return model;
    }
}