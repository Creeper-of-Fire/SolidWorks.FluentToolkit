using SolidWorks.Helpers;
using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Progress;

public class DrawFlange(ISldWorks sldWorks) : AbstractRun(sldWorks)
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
        Console.WriteLine("\n--- 开始建模: 带水线和螺栓孔的复杂法兰 ---");

        // --- 步骤 1: 定义几何尺寸 (单位: 毫米) ---
        // 法兰主体尺寸
        const double flangeRadiusR = 100.0; // R: 法兰外半径
        const double flangeThicknessC = 25.0; // C: 法兰盘厚度
        const double totalLengthH = 90.0; // H: 总长度
        const double neckRadiusR1 = 55.0; // R1: 颈部外半径
        const double boreRadiusR2 = 40.0; // R2: 内部通孔半径
        const double neckChamferAngle = 50.0; // 颈部斜角 (度)

        // 水线尺寸
        const double waterlineStartRadius = 45.0; // 水线起始半径
        const double waterlineEndRadius = 65.0; // 水线终止半径
        const double waterlineSpacing = 2.0; // 水线间距
        const double waterlineGrooveDepth = 0.5; // 水线凹槽深度
        const double grooveWidth = waterlineSpacing / 2.0; // 每个凹槽的宽度

        // 螺栓孔尺寸
        const double boltCircleRadius = 80; // 螺栓孔中心圆半径
        const double boltHoleRadius = 9.0; // 螺栓孔半径
        const int boltHoleCount = 8; // 螺栓孔数量

        Console.WriteLine("日志：所有几何参数已定义。");

        // --- 步骤 2: 使用旋转命令创建法兰主体 ---
        Console.WriteLine("\n--- 步骤 2: 创建旋转凸台基体 ---");

        // 基于尺寸计算轮廓顶点坐标 (全局3D坐标)
        double chamferDeltaY = neckRadiusR1 - boreRadiusR2;
        double chamferDeltaZ = chamferDeltaY * Math.Tan(neckChamferAngle * Math.PI / 180.0);

        var contourPoints = new Point3D[]
        {
            new(0, boreRadiusR2, 0.0), // p1: 内孔后边缘
            new(0, flangeRadiusR, 0.0), // p2: 外盘后边缘
            new(0, flangeRadiusR, flangeThicknessC), // p3: 外盘前边缘
            new(0, neckRadiusR1, flangeThicknessC), // p4: 颈部与盘连接处
            new(0, neckRadiusR1, totalLengthH - chamferDeltaZ), // p5: 颈部斜角开始处
            new(0, boreRadiusR2, totalLengthH) // p6: 颈部斜角结束处
        };
        Console.WriteLine("日志：已计算出主体截面轮廓的6个全局3D坐标点。");

        swModel.SelectByName("右视基准面", "PLANE") // 在右视基准面 (YZ平面) 上绘图
            .CreateRevolveBoss(sm =>
            {
                // 创建中心线作为旋转轴 (绕Z轴旋转)
                sm.CreateCenterLine(-10, 0, 0, totalLengthH + 10, 0, 0);
                // 使用 DrawContour 绘制整个封闭轮廓
                sm.DrawContour(contourPoints);
            }, out var revolveFeature)
            .ClearSelection();

        Console.WriteLine("日志：法兰主体旋转特征已成功创建！");


        // --- 步骤 3: 在法兰盘平面上创建水线凹槽 ---
        Console.WriteLine("\n--- 步骤 3: 循环创建水线凹槽 (旋转切除) ---");
        
        // 我们需要在与主体相同的平面上进行旋转切除
        swModel.SelectByName("右视基准面", "PLANE", append: false)
            .CreateRevolveCut(sm =>
            {
                Console.WriteLine("日志：正在创建旋转切除的中心线...");
                sm.CreateCenterLine(-10, 0, 0, totalLengthH + 10, 0, 0); // 同样的旋转轴
        
                int waterlineCount = (int)((waterlineEndRadius - waterlineStartRadius) / waterlineSpacing);
                for (int i = 0; i <= waterlineCount; i++)
                {
                    double currentRadius = waterlineStartRadius + i * waterlineSpacing;
                    Console.WriteLine($"日志：正在为第 {i + 1}/{waterlineCount + 1} 条水线绘制截面，半径 {currentRadius}mm");
        
                    // 凹槽从Z=0平面切入，深度为waterlineGrooveDepth
                    // 在“右视基准面”草图中，全局Z轴对应草图的-X轴
                    const double rectStartX_Sketch = 0; // 对应全局 Z=0
                    const double rectEndX_Sketch = -waterlineGrooveDepth; // 对应全局 Z=waterlineGrooveDepth
        
                    // Y坐标控制凹槽的径向位置
                    double rectY1 = currentRadius;
                    double rectY2 = rectY1 + grooveWidth;
        
                    sm.CreateCornerRectangle(rectStartX_Sketch, rectY1, 0, rectEndX_Sketch, rectY2, 0)
                        .AssertNotNull($"创建第{i + 1}个水线矩形失败");
                }
            }, out _);
        
        swModel.ClearSelection();
        Console.WriteLine("日志：所有水线凹槽已通过旋转切除创建。");
        
        
        // --- 步骤 4: 创建单个螺栓孔 (为阵列做准备) ---
        Console.WriteLine("\n--- 步骤 4: 创建单个螺栓孔 ---");

        // 选择法兰盘的大平面 (Z=0) 来绘制草图
        // 从一个已知在模型外部的点 (0, Y, -10)，沿着Z轴正方向发射一条射线。
        // 这条射线将会精确地“击中”并选中我们想要操作的法兰盘平面。
        swModel
            .SelectByRay(
                (0, (flangeRadiusR + boreRadiusR2) / 2.0, -10),
                (0, 0, 1)
            )
            .CreateCutThroughAll(
                sm => sm.CreateCircleByRadius(0, boltCircleRadius, 0, boltHoleRadius),
                out var boltHoleSeedFeature)
            .ClearSelect();

        Console.WriteLine("日志：单个螺栓孔（阵列源）切除成功。");

        // --- 步骤 5: 圆周阵列螺栓孔 ---
        Console.WriteLine("\n--- 步骤 5: 创建螺栓孔圆周阵列 ---");

        swModel
            // 选择内孔的圆柱面作为阵列中心轴 (最稳健的方法)
            .SelectByRay(
                (0, 0, totalLengthH / 2.0),
                (1, 0, 0),
                mark: 1
            )
            // 选择上一步创建的切除特征作为阵列源
            .SelectFeature(boltHoleSeedFeature, mark: 4)
            // 执行圆周阵列
            .CreateCircularPattern(boltHoleCount, out _)
            .ClearSelect();

        Console.WriteLine($"日志：成功创建了 {boltHoleCount} 个螺栓孔的圆周阵列。");

        // --- 步骤 6: 添加圆角和倒角，优化模型 ---
        Console.WriteLine("\n--- 步骤 6: 添加圆角和倒角，优化模型 ---");

        // --- 定义所有需要的几何体的“查找规则” ---
        var mainBody = new LazyRef<IBody2>(() =>
        {
            var part = (PartDoc)swModel;
            var bodies = (object[])part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            return bodies.Cast<IBody2>().FirstOrDefault().AssertNotNull("未能找到零件中的主实体。");
        });

        // 面(Face)的查找规则
        var neckPlaneFace = new LazyRef<IFace2>(() =>
            GeometryHelper.FindPlanarFace(mainBody.Value, new Point3D(0, 0, 1), new Point3D(0, 0, flangeThicknessC))
                .AssertNotNull("未能找到法兰盘平面。")
        );
        var outerCylinderFace = new LazyRef<IFace2>(() =>
            GeometryHelper.FindCylindricalFaceByRadius(mainBody.Value, flangeRadiusR)
                .AssertNotNull("未找到法兰外圆柱面。")
        );
        var neckCylinderFace = new LazyRef<IFace2>(() =>
            GeometryHelper.FindCylindricalFaceByRadius(mainBody.Value, neckRadiusR1)
                .AssertNotNull("未找到颈部圆柱面。")
        );
        var boreFace = new LazyRef<IFace2>(() =>
            GeometryHelper.FindCylindricalFaceByRadius(mainBody.Value, boreRadiusR2)
                .AssertNotNull("未能找到内孔圆柱面。")
        );
        var boreEndFace = new LazyRef<IFace2>(() =>
            GeometryHelper.FindConicalFaceByBoundaryCircle(mainBody.Value, boreRadiusR2, new Point3D(0, 0, totalLengthH))
                .AssertNotNull("未能找到内孔前端的锥形面。")
        );

        // 边(Edge)的查找规则 (这些规则依赖于上面的面规则)
        var outerChamferEdge = new LazyRef<IEdge>(() =>
            GeometryHelper.GetIntersectionEdge(outerCylinderFace.Value, neckPlaneFace.Value)
                .AssertNotNull("未能找到法兰外缘倒角边线。")
        );
        var neckFilletEdge = new LazyRef<IEdge>(() =>
            GeometryHelper.GetIntersectionEdge(neckCylinderFace.Value, neckPlaneFace.Value)
                .AssertNotNull("未能找到颈部圆角边线。")
        );
        var boreChamferEdge = new LazyRef<IEdge>(() =>
            GeometryHelper.GetIntersectionEdge(boreFace.Value, boreEndFace.Value)
                .AssertNotNull("未能找到内孔前端倒角边线。")
        );

        // 在颈部与法兰盘连接处添加一个平滑的圆角
        const double neckFilletRadius = 5.0;
        Console.WriteLine($"日志：正在颈部连接处添加 R{neckFilletRadius} 的圆角...");
        swModel.SelectObject(neckFilletEdge.Value)
            .CreateConstantRadiusFillet(neckFilletRadius, out _)
            .ClearSelect();
        Console.WriteLine("日志：颈部圆角创建成功。");

        // 在法兰外缘添加一个1mm*45°的倒角，防止刮手
        const double outerChamferDist = 1.0;
        Console.WriteLine($"日志：正在法兰外缘添加 {outerChamferDist}x45° 的倒角...");
        swModel.SelectObject(outerChamferEdge.Value)
            .CreateAngleDistanceChamfer(outerChamferDist, 45.0, out _)
            .ClearSelect();
        Console.WriteLine("日志：法兰外缘倒角创建成功。");

        // 在内孔前端添加一个1mm*45°的倒角，方便装配
        const double boreChamferDist = 1.0;
        Console.WriteLine($"日志：正在内孔前端添加 {boreChamferDist}x45° 的倒角...");
        swModel.SelectObject(boreChamferEdge.Value)
            .CreateAngleDistanceChamfer(boreChamferDist, 45.0, out _)
            .ClearSelect();
        Console.WriteLine("日志：内孔前端倒角创建成功。");

        // --- 模型收尾 ---
        swModel.ViewZoomtofit2();
        swModel.ShowNamedView2("*等轴测", -1);
        Console.WriteLine("\n日志：带水线和螺栓孔的复杂法兰建模完成！");
    }
}