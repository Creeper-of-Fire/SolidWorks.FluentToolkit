using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Progress;

public class DrawFlange(SldWorks sldWorks) : AbstractRun(sldWorks)
{
    public override void Run()
    {
        // 新建一个零件
        ModelDoc2 swModel = (ModelDoc2)this.SldWorks.NewPart().AssertNotNull("未能创建新零件。");

        Console.WriteLine($"日志：新零件 '{swModel.GetTitle()}' 已创建。");

        // --- 核心建模代码开始 ---

        // 为了方便，我们获取几个常用的“管理器”对象
        FeatureManager swFeatureManager = swModel.FeatureManager;
        SketchManager swSketchManager = swModel.SketchManager;
        ModelDocExtension swModelDocExt = swModel.Extension;
        SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;

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

        swModelDocExt.SelectByID2("前视基准面", "PLANE", 0, 0, 0, false, 0, null, 0)
            .AssertTrue("选择“前视基准面”失败！请检查 SOLIDWORKS 语言或模板。");

        swSketchManager.InsertSketch(true);
        // 检查草图是否成功创建
        swSketchManager.ActiveSketch.AssertNotNull("进入草图模式失败！");

        swSketchManager.CreateCircleByRadius(0, 0, 0, flangeOuterDiameter / 2.0);
        Console.WriteLine("日志：绘制了法兰外轮廓。");
        swSketchManager.InsertSketch(true);


        Feature flangeBodyFeature = swFeatureManager.FeatureExtrusion3(
                true, false, false, (int)swEndConditions_e.swEndCondBlind, 0, flangeThickness, 0,
                false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false)
            .AssertNotNull("创建法兰本体拉伸特征失败！");

        Console.WriteLine($"日志：法兰本体拉伸成功，厚度 {flangeThickness * 1000}mm。");
        swModel.ClearSelection2(true);


        // --- 步骤 2: 创建中心通孔 (拉伸切除) ---
        Console.WriteLine("\n--- 步骤 2: 创建中心通孔 ---");

        // 【校验 2.1】检查法兰正面选择是否成功
        swModelDocExt.SelectByID2("", "FACE", 0, 0, flangeThickness, false, 0, null, 0)
            .AssertTrue("选择法兰盘正面用于中心孔草图失败！");

        swSketchManager.InsertSketch(true);

        // 检查草图是否成功创建
        swSketchManager.ActiveSketch.AssertNotNull("在法兰盘正面为中心孔创建草图失败！");

        swSketchManager.CreateCircleByRadius(0, 0, 0, pipeInnerDiameter / 2.0);
        Console.WriteLine("日志：绘制了中心孔轮廓。");
        swSketchManager.InsertSketch(true);

        // 【校验 2.3】检查拉伸切除特征是否成功创建
        Feature centerCutFeature = swFeatureManager.FeatureCut4(
                true, false, false, (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll,
                0, 0, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false,
                false)
            .AssertNotNull("创建中心通孔拉伸切除特征失败！");

        Console.WriteLine("日志：中心通孔切除成功。");
        swModel.ClearSelection2(true);


        // --- 步骤 3: 创建单个螺栓孔 (为阵列做准备) ---
        Console.WriteLine("\n--- 步骤 3: 创建单个螺栓孔 ---");

        double selectionPointX = (pipeInnerDiameter + flangeOuterDiameter) / 4.0;
        // 【校验 3.1】检查法兰正面选择是否成功 (使用安全坐标点)
        swModelDocExt.SelectByID2("", "FACE", selectionPointX, 0, flangeThickness, false, 0, null, 0)
            .AssertTrue("选择法兰盘正面用于螺栓孔草图失败！");

        swSketchManager.InsertSketch(true);

        // 【校验 3.2】检查草图是否成功创建
        swSketchManager.ActiveSketch.AssertNotNull("严重错误：在法兰盘正面为螺栓孔创建草图失败！");

        swSketchManager.CreateCircleByRadius(0, boltCircleDiameter / 2.0, 0, boltHoleDiameter / 2.0);
        Console.WriteLine("日志：绘制了第一个螺栓孔轮廓。");
        swSketchManager.InsertSketch(true);

        // 【校验 3.3】检查拉伸切除特征是否成功创建
        Feature boltHoleSeedFeature = swFeatureManager.FeatureCut4(
                true, false, false, (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll,
                0, 0, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false,
                false)
            .AssertNotNull("创建单个螺栓孔拉伸切除特征失败！");

        Console.WriteLine("日志：单个螺栓孔（阵列源）切除成功。");
        swModel.ClearSelection2(true);


        // --- 步骤 4: 圆周阵列螺栓孔 ---
        Console.WriteLine("\n--- 步骤 4: 创建螺栓孔圆周阵列 ---");

        // 1. 选择阵列轴：选择中心孔的圆形边线。这非常稳定。
        //    我们使用一个在边线上的精确坐标 (X=内半径, Y=0, Z=厚度) 来选中它。
        //    SelectByID2 的 Mark 参数对于阵列轴是 1。
        swModelDocExt.SelectByID2("", "EDGE", pipeInnerDiameter / 2.0, 0, flangeThickness, false, 1, null, 0)
            .AssertTrue("选择阵列旋转轴（中心孔上边缘线）失败！");

        Console.WriteLine("日志：已选择阵列旋转轴。");

        // 2. 选择要阵列的特征本身。
        //    注意 SelectByID2 的第二个参数 Append 设为 true，这样才能将特征添加到已选的轴上。
        //    Mark 参数对于要阵列的特征是 4。
        swModelDocExt.SelectByID2(boltHoleSeedFeature.Name, "BODYFEATURE", 0, 0, 0, true, 4, null, 0)
            .AssertTrue($"选择要阵列的特征 '{boltHoleSeedFeature.Name}' 失败！");

        Console.WriteLine($"日志：已选择要阵列的特征 '{boltHoleSeedFeature.Name}'。");

        // 3. 创建圆周阵列特征。
        //    GeometryPattern 设置为 false，表示我们正在进行特征阵列。
        Feature circularPatternFeature = swFeatureManager.FeatureCircularPattern4(
                boltHoleCount, // 参数 1: Number - 实例总数
                (360.0 * Math.PI / 180.0), // 参数 2: Spacing - 阵列总角度 (360度，需转为弧度)
                false, // 参数 3: FlipDirection - 反向
                "", // 参数 4: DName - 特征尺寸名称 (通常为空)
                false, // 参数 5: GeometryPattern - 设为 false，执行特征阵列
                true, // 参数 6: EqualSpacing - 均布
                false // 参数 7: VaryInstance - 实例随形变化
            )
            .AssertNotNull("创建圆周阵列失败！");

        Console.WriteLine($"日志：成功创建了 {boltHoleCount} 个螺栓孔的圆周阵列。");
        swModel.ClearSelection2(true);


        // --- 模型收尾 ---
        swModel.ViewZoomtofit2();
        swModel.ShowNamedView2("*等轴测", -1);

        Console.WriteLine("\n日志：法兰盘建模完成！");
    }
}