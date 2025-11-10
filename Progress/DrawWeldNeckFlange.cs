using SolidWorks.Helpers;
using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Progress;

/// <summary>
/// 根据标准图纸，创建一个带颈平焊法兰 (Weld Neck Flange)。
/// </summary>
public class DrawWeldNeckFlange(ISldWorks sldWorks) : AbstractRun(sldWorks)
{
    public override void Run()
    {
        // --- 准备工作: 新建零件并设置单位 ---
        ModelDoc2 swModel = (ModelDoc2)this.SldWorks.NewPart().AssertNotNull("未能创建新零件。");
        Console.WriteLine($"日志：新零件 '{swModel.GetTitle()}' 已创建。");
        swModel.Extension.SetUserPreferenceInteger(
            (int)swUserPreferenceIntegerValue_e.swUnitSystem, 0, (int)swUnitSystem_e.swUnitSystem_MMGS);
        Console.WriteLine($"日志：文档单位已设置为 MMGS (毫米、克、秒)。");

        // --- 核心建模代码开始 ---
        Console.WriteLine("\n--- 开始建模: 带颈平焊法兰 ---");

        // --- 步骤 1: 定义几何尺寸 (单位: 毫米) ---
        const double flangeDiameterD = 200.0;    // D: 法兰外径
        const double flangeThicknessC = 25.0;     // C: 法兰盘厚度
        const double totalLengthH = 90.0;         // H: 总长度
        const double neckDiameterA1 = 110.0;      // A1: 颈部外径
        const double boreDiameterB1 = 80.0;       // B1: 内部通孔直径
        const double neckChamferAngle = 50.0;     // 颈部斜角 (度)

        Console.WriteLine("日志：法兰几何参数已定义。");

        // --- 步骤 2: 基于尺寸计算轮廓顶点坐标 (使用全局3D坐标) ---
        // 我们计划在 "右视基准面" (YZ平面) 上绘图，所以所有点的 X 坐标都为 0。
        // Z 坐标对应草图的 X 轴，Y 坐标对应草图的 Y 轴。
        double radiusD = flangeDiameterD / 2.0;
        double radiusA1 = neckDiameterA1 / 2.0;
        double radiusB1 = boreDiameterB1 / 2.0;
        double chamferDeltaY = radiusA1 - radiusB1;
        double chamferDeltaZ = chamferDeltaY * Math.Tan(neckChamferAngle * Math.PI / 180.0);
        
        // 定义全局坐标点 (X=0, Y, Z)
        var contourPoints = new Point3D[]
        {
            new (0, radiusB1, 0.0),                      // p1: 内孔后边缘
            new (0, radiusD, 0.0),                       // p2: 外盘后边缘
            new (0, radiusD, flangeThicknessC),          // p3: 外盘前边缘
            new (0, radiusA1, flangeThicknessC),         // p4: 颈部与盘连接处
            new (0, radiusA1, totalLengthH - chamferDeltaZ), // p5: 颈部斜角开始处
            new (0, radiusB1, totalLengthH)              // p6: 颈部斜角结束处
        };
        Console.WriteLine("日志：已根据参数计算出截面轮廓的6个全局3D坐标点。");

        // --- 步骤 3: 使用旋转命令创建法兰主体 ---
        Console.WriteLine("\n--- 步骤 3: 创建旋转凸台基体 ---");

        swModel.SelectByName("右视基准面", "PLANE") // 在右视基准面 (YZ平面) 上绘图
            .CreateRevolveBoss(sm =>
            {
                Console.WriteLine("日志：进入草图，开始绘制轮廓...");

                // 1. 创建中心线作为旋转轴 (必须是第一条构造线)
                // 我们绕Z轴旋转，在YZ平面草图中，Z轴是X轴
                sm.CreateCenterLine(-10, 0, 0, totalLengthH + 10, 0, 0);
                Console.WriteLine("日志：已创建旋转中心线。");

                // 2. 使用 DrawContour 绘制整个封闭轮廓
                sm.DrawContour(contourPoints);

                Console.WriteLine("日志：截面轮廓绘制完成。");
            },out _)
            .ClearSelection();

        Console.WriteLine("日志：旋转特征已成功创建！");

        // --- 模型收尾 ---
        swModel.ViewZoomtofit2();
        swModel.ShowNamedView2("*等轴测", -1);
        Console.WriteLine("\n日志：带颈平焊法兰建模完成！");
    }
}