using SolidWorks.Helpers;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Progress;

public class DrawComplexFlange(ISldWorks sldWorks) : AbstractRun(sldWorks)
{
    public override void Run()
    {
        // 新建一个零件
        ModelDoc2 swModel = (ModelDoc2)this.SldWorks.NewPart().AssertNotNull("未能创建新零件。");
        Console.WriteLine($"日志：新零件 '{swModel.GetTitle()}' 已创建。");

        var modelDocExt = swModel.Extension;
        modelDocExt.SetUserPreferenceInteger(
            (int)swUserPreferenceIntegerValue_e.swUnitSystem,
            0,
            (int)swUnitSystem_e.swUnitSystem_MMGS);
        Console.WriteLine($"日志：文档单位已设置为 MMGS (毫米、克、秒)。");

        // --- 核心建模代码开始 ---

        // --- 定义几何尺寸 (单位: 毫米) ---
        const double flangeOuterDiameter = 150; // 法兰外径 150mm
        const double flangeThickness = 25; // 法兰厚度 25mm
        const double pipeInnerDiameter = 60; // 中心孔径 60mm
        // 水线参数
        const double waterlineStartDiameter = 65; // 水线起始直径
        const double waterlineEndDiameter = 80; // 水线终止直径
        const double waterlineSpacing = 2; // 水线间距 2mm
        const double waterlineGrooveDepth = 0.5; // 水线凹槽深度 1.0mm
        const double grooveWidth = waterlineSpacing / 2.0; // 每个凹槽的宽度

        Console.WriteLine("日志：复杂法兰尺寸参数已定义。");

        // --- 步骤 1: 创建法兰基体 (临时方案，后续将替换为旋转) ---
        Console.WriteLine("\n--- 步骤 1: 创建法兰基体 ---");

        // 步骤 1.1: 创建法兰盘本体
        swModel.SelectByName("前视基准面", "PLANE")
            .CreateBossExtrusion(flangeThickness,
                sm => sm.CreateCircleByRadius(0, 0, 0, flangeOuterDiameter / 2.0),out _)
            .ClearSelect(); // 操作完成后立即清空选择，养成好习惯

        Console.WriteLine("日志：法兰本体拉伸成功。");

        // 步骤 1.2: 创建中心通孔
        // 我们需要在新创建的法兰面上选择一个点来定义草图平面
        double selectionPointX = (pipeInnerDiameter + flangeOuterDiameter) / 4.0; // 在内圆和外圆之间选一个点

        swModel.SelectFaceByPoint(selectionPointX, 0, flangeThickness) // 在Z=flangeThickness的平面上选择一个面
            .CreateCutThroughAll(sm => sm.CreateCircleByRadius(0, 0, 0, pipeInnerDiameter / 2.0),out _)
            .ClearSelect();

        Console.WriteLine("日志：中心通孔切除成功，法兰基体完成。");

        // --- 步骤 2: 创建中心旋转轴 ---
        Console.WriteLine("\n--- 步骤 2: 创建中心旋转轴 ---");

        // 为了进行旋转操作，我们需要一个明确的、作为特征存在的旋转轴。
        // 通过相交“前视基准面”和“上视基准面”来创建这个轴。
        swModel.CreateReferenceAxis(out var axisFeature).ClearSelect();
        Console.WriteLine($"日志：中心旋转轴 '{axisFeature.Name}' 已创建。");

        // --- 步骤 3: 循环创建每一个水线凹槽 ---
        Console.WriteLine("\n--- 步骤 2: 循环创建水线凹槽 ---");

        int waterlineCount = (int)((waterlineEndDiameter - waterlineStartDiameter) / 2 / waterlineSpacing);


        swModel.SelectObject(axisFeature, append: false)
            .SelectByName("右视基准面", "PLANE", append: true);

        // Z 坐标控制凹槽的【轴向】位置 (深度)。
        const double rectZ1 = flangeThickness; // 从法兰顶面开始
        const double rectZ2 = rectZ1 - waterlineGrooveDepth; // 切入指定深度
        
        swModel.CreateRevolveCut(sm =>
        {
            Console.WriteLine("日志：正在创建旋转切除的中心线作为旋转轴...");
            sm.CreateCenterLine(-flangeOuterDiameter, 0, 0, flangeOuterDiameter, 0, 0);

            for (int i = 0; i <= waterlineCount; i++)
            {
                double currentDiameter = waterlineStartDiameter;
                double currentRadius = currentDiameter / 2.0 + i * waterlineSpacing;

                Console.WriteLine($"日志：正在创建第 {i + 1}/{waterlineCount + 1} 条水线，半径 {currentRadius} ...");

                // Y 坐标控制凹槽的【径向】位置。
                double rectY1 = currentRadius;
                double rectY2 = rectY1 + grooveWidth;
                
                // 在“右视基准面”上绘图时，草图的X轴指向全局的负Z轴。
                // 草图的Y轴指向全局的正Y轴。
                // 要在全局Z=a的位置创建点，需要提供草图X坐标为-a。
                object rectSegments = sm.CreateCornerRectangle(-rectZ1, rectY1, 0, -rectZ2, rectY2, 0);
                rectSegments.AssertNotNull("创建失败");
            }
        }, out _); // 执行旋转切除

        swModel.ClearSelect(); // 清理选择，准备下一次循环


        Console.WriteLine("日志：所有水线凹槽已创建。");

        // --- 模型收尾 ---
        swModel.ViewZoomtofit2();
        swModel.ShowNamedView2("*等轴测", -1);
        Console.WriteLine("\n日志：带水线的法兰 (阶段一) 建模完成！");
    }
}