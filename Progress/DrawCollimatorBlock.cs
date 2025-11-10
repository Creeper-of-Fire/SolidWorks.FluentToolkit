using SolidWorks.Helpers;
using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Tools;
using SolidWorks.Utils;

namespace SolidWorks.Progress;

public class DrawCollimatorBlock(ISldWorks sldWorks) : AbstractRun(sldWorks)
{
    public override void Run()
    {
        // --- 准备工作: 新建零件并设置单位为米 ---
        ModelDoc2 swModel = (ModelDoc2)this.SldWorks.NewPart().AssertNotNull("未能创建新零件。");
        Console.WriteLine($"日志：新零件 '{swModel.GetTitle()}' 已创建。");
        // 对于大型结构，使用米作为基本单位更合适
        swModel.Extension.SetUserPreferenceInteger(
            (int)swUserPreferenceIntegerValue_e.swUnitSystem, 0, (int)swUnitSystem_e.swUnitSystem_MKS);
        Console.WriteLine($"日志：文档单位已设置为 MKS (米、千克、秒)。");

        // --- 核心建模代码开始 ---
        Console.WriteLine("\n--- 开始建模: 参数化扇形准直孔阵列 ---");

        // --- 步骤 1: 定义几何尺寸 (单位: 米) ---
        const double blockWidth = 1.0; // 宽度 (X方向)
        const double blockHeight = 2.0; // 高度 (Y方向)
        const double blockDepth = 4.0; // 深度 (Z方向)

        // 准直孔参数
        const double holeDiameter = 0.05; // 孔径 ɸ50mm = 0.05m
        const int holesInX = 5; // X方向的孔数量
        const int holesInY = 10; // Y方向的孔数量

        // 定义扇形焦点 (所有孔都指向这个点)
        // 将焦点设置在实体后方 10 米处，位于中心线上
        var focusPoint = new Point3D(0, 0, blockDepth + 10.0);

        Console.WriteLine("日志：所有几何参数已定义。");
        Console.WriteLine($"日志：准直孔将汇聚于焦点: ({focusPoint.X}, {focusPoint.Y}, {focusPoint.Z})");

        // --- 步骤 2: 创建基础实体块 ---
        Console.WriteLine("\n--- 步骤 2: 创建基础实体块 ---");
        swModel.SelectByName("前视基准面", "PLANE")
            .CreateBossExtrusion(blockDepth, sm =>
            {
                // 定义矩形的四个角点 (全局3D坐标)
                var corner1 = new Point3D(-blockWidth / 2, blockHeight / 2, 0); // 左上
                var corner2 = new Point3D(blockWidth / 2, blockHeight / 2, 0); // 右上
                var corner3 = new Point3D(blockWidth / 2, -blockHeight / 2, 0); // 右下
                var corner4 = new Point3D(-blockWidth / 2, -blockHeight / 2, 0); // 左下

                // 使用我们可靠的 DrawContour 辅助函数来绘制轮廓
                sm.DrawContour([corner1, corner2, corner3, corner4]);
            }, out var baseFeature)
            .ClearSelect();

        var baseBody = baseFeature.GetFirstBody();
        Console.WriteLine($"日志：已创建 {blockWidth}m x {blockHeight}m x {blockDepth}m 的基础实体块。");

        // --- 步骤 3: 为矩形区域生成入口点列表 ---
        Console.WriteLine("\n--- 步骤 3: 正在为矩形表面计算入口点坐标... ---");
        var entryPoints = new List<Point3D>();
        double spacingX = blockWidth / (holesInX + 1);
        double spacingY = blockHeight / (holesInY + 1);

        for (int i = 1; i <= holesInY; i++)
        {
            for (int j = 1; j <= holesInX; j++)
            {
                double startX = -blockWidth / 2 + j * spacingX;
                double startY = -blockHeight / 2 + i * spacingY;
                entryPoints.Add(new Point3D(startX, startY, 0));
            }
        }

        Console.WriteLine($"日志：已生成 {entryPoints.Count} 个入口点。");

        // --- 步骤 4: 调用通用的阵列创建函数 ---
        Console.WriteLine("\n--- 步骤 4: 调用通用函数创建扇形孔阵列... ---");
        swModel.CreateConvergingSweptCuts(
            entryPoints,
            focusPoint,
            holeDiameter,
            out var tempSketches // 函数返回所有需要隐藏的草图
        );

        // --- 步骤 5: 收尾工作 ---
        Console.WriteLine("\n--- 步骤 5: 模型收尾 ---");

        // 隐藏所有的临时路径草图
        foreach (var sketchFeat in tempSketches)
        {
            swModel.SelectObject(sketchFeat)
                .BlankSketch(); // BlankSketch 用于隐藏草图
        }

        swModel.ClearSelect();
        Console.WriteLine($"日志：已隐藏 {tempSketches.Count} 个临时路径草图。");

        swModel.ViewZoomtofit2();
        swModel.ShowNamedView2("*等轴测", -1);
        Console.WriteLine("\n日志：参数化扇形准直孔阵列建模完成！");
    }
}