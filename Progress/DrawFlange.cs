using SolidWorks.Helpers;
using SolidWorks.Interop.sldworks;
using SolidWorks.Utils;
using Xarial.XCad.SolidWorks;

namespace SolidWorks.Progress;

public class DrawFlange(ISwApplication app) : AbstractRun(app)
{
    public override void Run()
    {
        // 新建一个零件
        ModelDoc2 swModel = (ModelDoc2)this.SldWorks.NewPart().AssertNotNull("未能创建新零件。");

        Console.WriteLine($"日志：新零件 '{swModel.GetTitle()}' 已创建。");

        // --- 核心建模代码开始 ---

        // --- 定义法兰盘的几何尺寸 (单位: 米) ---
        double flangeOuterDiameter = 0.150; // 法兰外径 150mm
        double flangeThickness = 0.020; // 法兰厚度 20mm
        double pipeInnerDiameter = 0.060; // 中心孔径 (通管) 60mm
        double boltCircleDiameter = 0.110; // 螺栓孔中心圆直径 110mm
        double boltHoleDiameter = 0.015; // 螺栓孔直径 15mm
        int boltHoleCount = 6; // 螺栓孔数量 6个

        Console.WriteLine("日志：法兰盘尺寸参数已定义。");

        // --- 步骤 1: 创建法兰盘本体 (拉伸一个圆盘) ---
        Console.WriteLine("\n--- 步骤 1: 创建法兰盘本体 ---");

        swModel.SelectByName("前视基准面", "PLANE")
            .CreateBossExtrusion(flangeThickness,
                sketchManager => sketchManager.CreateCircleByRadius(0, 0, 0, flangeOuterDiameter / 2.0))
            .ClearSelect();

        Console.WriteLine($"日志：法兰本体拉伸成功，厚度 {flangeThickness * 1000}mm。");


        // --- 步骤 2: 创建中心通孔 (拉伸切除) ---
        Console.WriteLine("\n--- 步骤 2: 创建中心通孔 ---");

        swModel.SelectFaceByPoint(0, 0, flangeThickness)
            .CreateCutThroughAll(sketchManager => sketchManager.CreateCircleByRadius(0, 0, 0, pipeInnerDiameter / 2.0))
            .ClearSelect();

        Console.WriteLine("日志：中心通孔切除成功。");

        // --- 步骤 3: 创建单个螺栓孔 (为阵列做准备) ---
        Console.WriteLine("\n--- 步骤 3: 创建单个螺栓孔 ---");

        double selectionPointX = (pipeInnerDiameter + flangeOuterDiameter) / 4.0;

        swModel.SelectFaceByPoint(selectionPointX, 0, flangeThickness)
            .CreateCutThroughAll(
                sketchManager => sketchManager.CreateCircleByRadius(0, boltCircleDiameter / 2.0, 0, boltHoleDiameter / 2.0),
                out var boltHoleSeedFeature)
            .ClearSelect();

        Console.WriteLine("日志：单个螺栓孔（阵列源）切除成功。");


        // --- 步骤 4: 圆周阵列螺栓孔 ---
        Console.WriteLine("\n--- 步骤 4: 创建螺栓孔圆周阵列 ---");

        swModel.SelectEdgeByPoint(pipeInnerDiameter / 2.0, 0, flangeThickness, mark: 1) // 显式选择阵列轴
            .SelectFeature(boltHoleSeedFeature, mark: 4) // 显式选择要阵列的特征
            .CreateCircularPattern(boltHoleCount, out _) // 执行阵列
            .ClearSelect();

        Console.WriteLine($"日志：成功创建了 {boltHoleCount} 个螺栓孔的圆周阵列。");


        // --- 模型收尾 ---
        swModel.ViewZoomtofit2();
        swModel.ShowNamedView2("*等轴测", -1);

        Console.WriteLine("\n日志：法兰盘建模完成！");
    }
}