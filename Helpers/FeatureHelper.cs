using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Helpers;

/// <summary>
/// 为 IModelDoc2 提供创建Feature的扩展方法，以简化常用的建模操作。
/// </summary>
public static class FeatureHelper
{
    private static SldWorks SldWorks { get; } = SolidWorksConnector.GetSldWorksApp();

    /// <summary>
    /// (通用辅助方法) 执行一个会创建新特征的操作，并通过“差集法”安全地返回新创建的特征。
    /// 这种方法是语言无关且非常可靠的。
    /// </summary>
    /// <param name="featureManager">要进行操作的 FeatureManager。</param>
    /// <param name="createAction">一个委托，包含了实际创建特征的操作，例如 () => model.InsertHelix(...)</param>
    /// <returns>新创建的 Feature 对象。</returns>
    /// <exception cref="InvalidOperationException">如果操作后未能找到任何新特征，则抛出异常。</exception>
    public static Feature ExecuteAndGetNewFeature(IFeatureManager featureManager, Action createAction)
    {
        // 步骤 1: 操作前，获取所有特征的名称
        var featuresBefore = (object[])featureManager.GetFeatures(false) ?? Array.Empty<object>();
        var namesBefore = featuresBefore.Cast<Feature>().Select(f => f.Name).ToHashSet();

        // 步骤 2: 执行传入的创建操作
        createAction();

        // 步骤 3: 操作后，再次获取所有特征
        object[] featuresAfter = (object[])featureManager.GetFeatures(false) ?? [];

        // 步骤 4: 找出新特征
        var newFeature = featuresAfter
            .Cast<Feature>()
            .FirstOrDefault(f => !namesBefore.Contains(f.Name));

        // 步骤 5: 健壮性检查
        return newFeature.AssertNotNull("未能通过'差集法'找到新创建的特征。操作可能未成功创建任何新特征。");
    }

    /// <summary>
    /// 一个通用的私有辅助方法，用于在当前选定的平面上进入草图、执行绘图动作、然后退出草图。
    /// </summary>
    private static IModelDoc2 Sketch(this IModelDoc2 model, Action<ISketchManager> sketchAction, bool autoExit = true)
    {
        var sketchManager = model.SketchManager;

        if (sketchManager.ActiveSketch != null)
        {
            throw new InvalidOperationException("无法在已激活的草图环境中启动新的草图操作。请在调用此方法前退出当前草图。");
        }

        sketchManager.InsertSketch(true); // 进入草图

        if (sketchManager.ActiveSketch == null)
        {
            throw new InvalidOperationException("无法进入草图模式。请检查当前模型是否已打开草图。");
        }

        sketchAction(sketchManager); // 执行绘图动作

        if (autoExit)
        {
            sketchManager.InsertSketch(true); // 退出草图
            if (sketchManager.ActiveSketch != null)
            {
                throw new InvalidOperationException("无法退出草图模式。请检查当前模型是否已退出草图。");
            }
        }

        return model;
    }

    /// <summary>
    /// 在当前选定的基准面或平面上创建一个拉伸凸台。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="depth">拉伸的深度。</param>
    /// <param name="sketchAction">一个委托，定义了如何在草图中绘制几何图形。例如：sm => sm.CreateCircleByRadius(...)</param>
    /// <param name="createdFeature">创建的拉伸特征。</param>
    /// <returns></returns>
    public static IModelDoc2 CreateBossExtrusion(
        this IModelDoc2 model,
        double depth,
        Action<ISketchManager> sketchAction,
        out Feature createdFeature)
    {
        var featureManager = model.FeatureManager;

        // 核心流程：进入草图 -> 执行绘图动作 -> 退出草图
        model.Sketch(sketchAction);

        // 使用预设的、最常用的参数调用原生 API
        createdFeature = featureManager.FeatureExtrusion3(
                true, false, false, (int)swEndConditions_e.swEndCondBlind, 0, depth, 0,
                false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false)
            .AssertNotNull("创建拉伸凸台特征失败！");

        return model;
    }

    /// <summary>
    /// 在当前选定的平面上创建一个“完全贯穿”的拉伸切除特征。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="sketchAction">一个委托，定义了如何在草图中绘制几何图形。</param>
    /// <param name="createdFeature">创建的切除特征。</param>
    /// <returns></returns>
    public static IModelDoc2 CreateCutThroughAll(
        this IModelDoc2 model,
        Action<ISketchManager> sketchAction,
        out Feature createdFeature)
    {
        var featureManager = model.FeatureManager;

        model.Sketch(sketchAction);

        // 使用预设的“完全贯穿”参数调用原生 API
        createdFeature = featureManager.FeatureCut4(
                true, false, false,
                (int)swEndConditions_e.swEndCondThroughAll,
                (int)swEndConditions_e.swEndCondThroughAll,
                0, 0, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false,
                false)
            .AssertNotNull("创建拉伸切除特征失败！");

        return model;
    }

    /// <summary>
    /// 通过相交两个基准面（前视和上视）来创建一条中心基准轴。
    /// 这对于后续的旋转操作非常有用。
    /// TODO 改成可选轴方向的
    /// 
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="createdFeature">创建的基准轴特征。</param>
    public static IModelDoc2 CreateReferenceAxis(this IModelDoc2 model, out Feature createdFeature)
    {
        createdFeature = ExecuteAndGetNewFeature(model.FeatureManager, () =>
        {
            model
                .SelectByName("上视基准面", "PLANE", append: false)
                .SelectByName("右视基准面", "PLANE", append: true);
            model.InsertAxis2(true).AssertTrue("通过基准面相交创建基准轴失败！");
        });

        return model;
    }

    /// <summary>
    /// 在当前选定的基准面上，围绕第一条构造线（中心线）进行旋转操作。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="sketchAction">一个委托，定义了如何在草图中绘制几何图形。第一条创建的构造线将被用作旋转轴。</param>
    /// <param name="isCut">用于区分是凸台/基体还是切除</param>
    /// <param name="createdFeature">创建的旋转操作特征。</param>
    public static IModelDoc2 CreateRevolveFeature(
        this IModelDoc2 model,
        Action<ISketchManager> sketchAction,
        bool isCut,
        out Feature createdFeature)
    {
        var featureManager = model.FeatureManager;

        model.Sketch(sketchAction);

        // FeatureRevolve2 用于创建旋转凸台和旋转切除
        // 最后一个参数 'isCut' 设为 true，表示执行切除
        createdFeature = featureManager.FeatureRevolve2(
            true, // SingleDir: true 表示单向旋转
            true, // IsSolid: true 表示这是一个实体特征
            false, // IsThin: false 表示不是薄壁特征
            isCut, // IsCut
            false, // ReverseDir: false 表示不反转方向
            false, // BothDirectionUpToSameEntity (通常为 false)
            (int)swEndConditions_e.swEndCondBlind, // Dir1Type: 方向一终止条件 = 给定角度
            (int)swEndConditions_e.swEndCondBlind, // Dir2Type: 方向二终止条件 (因单向而无效)
            360.0 * Math.PI / 180.0, // Dir1Angle: 旋转 360 度
            0, // Dir2Angle (无效)
            false, // OffsetReverse1
            false, // OffsetReverse2
            0, // OffsetDistance1
            0, // OffsetDistance2
            0, // ThinType (无效)
            0, // ThinThickness1 (无效)
            0, // ThinThickness2 (无效)
            true, // Merge: true 表示与现有实体合并结果 (对于切除，这意味着从实体中减去)
            false, // UseFeatScope: false 表示影响所有实体
            true // UseAutoSelect: true 表示自动选择受影响的实体
        ).AssertNotNull("创建旋转切除特征失败！");

        return model;
    }

    /// <inheritdoc cref="CreateRevolveFeature(IModelDoc2,Action{ISketchManager},bool,out Feature)"/>
    /// <remarks>在当前选定的基准面上，围绕第一条构造线（中心线）进行旋转凸台/基体操作。</remarks>
    public static IModelDoc2 CreateRevolveBoss(this IModelDoc2 model, Action<ISketchManager> sketchAction, out Feature createdFeature)
    {
        return CreateRevolveFeature(model, sketchAction, isCut: false, out createdFeature);
    }

    /// <inheritdoc cref="CreateRevolveFeature(IModelDoc2,Action{ISketchManager},bool,out Feature)"/>
    /// <remarks>在当前选定的基准面上，围绕第一条构造线（中心线）进行旋转切除。</remarks>
    public static IModelDoc2 CreateRevolveCut(this IModelDoc2 model, Action<ISketchManager> sketchAction, out Feature createdFeature)
    {
        return CreateRevolveFeature(model, sketchAction, isCut: true, out createdFeature);
    }

    /// <summary>
    /// 在当前选定的基准面或平面上创建一个指定深度的拉伸切除。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="depth">切除的深度（单位：米）。</param>
    /// <param name="sketchAction">一个委托，定义了如何在草图中绘制几何图形。</param>
    /// <param name="createdFeature">创建的切除特征。</param>
    public static IModelDoc2 CreateCutExtrusion(
        this IModelDoc2 model,
        double depth,
        Action<ISketchManager> sketchAction,
        out Feature createdFeature)
    {
        var featureManager = model.FeatureManager;

        model.Sketch(sketchAction);

        // 使用 FeatureCut4，但指定 EndCondition 为 Blind (给定深度)
        createdFeature = featureManager.FeatureCut4(
                true, // Sd: true 表示这是一个标准的、非钣金的切除
                false, // Flip: false 表示不反转切除侧 (切除草图内部)
                false, // Dir: false 表示不启用第二方向 (单向切除)
                (int)swEndConditions_e.swEndCondBlind, // T1: 方向一的终止条件 = 给定深度
                (int)swEndConditions_e.swEndCondBlind, // T2: 方向二的终止条件 (由于Dir=false，此项无影响，设为Blind最安全)
                depth, // D1: 方向一的深度
                0, // D2: 方向二的深度 (由于Dir=false，此项无影响，设为0)
                false, // Dchk1: 方向一不启用拔模
                false, // Dchk2: 方向二不启用拔模
                false, // Ddir1: 方向一拔模向外 (无影响)
                false, // Ddir2: 方向二拔模向外 (无影响)
                0, // Dang1: 方向一拔模角度 (无影响)
                0, // Dang2: 方向二拔模角度 (无影响)
                false, // OffsetReverse1
                false, // OffsetReverse2
                false, // TranslateSurface1
                false, // TranslateSurface2
                false, // NormalCut: false 表示不是钣金的“垂直于折弯的切除”
                false, // UseFeatScope: false 表示特征影响所有实体
                true, // UseAutoSelect: true 表示自动选择所有受影响的实体
                false, // AssemblyFeatureScope
                false, // AutoSelectComponents
                false, // PropagateFeatureToParts
                0, // T0: 起始偏移条件 (0 = 从草图平面开始)
                0, // StartOffset: 起始偏移距离
                false, // FlipStartOffset
                false // OptimizeGeometry
            )
            .AssertNotNull("创建拉伸切除特征失败！");

        return model;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="model"></param>
    /// <param name="instanceCount"></param>
    /// <param name="createdFeature"></param>
    /// <returns></returns>
    public static IModelDoc2 CreateCircularPattern(this IModelDoc2 model, int instanceCount, out Feature createdFeature)
    {
        createdFeature = model.FeatureManager.FeatureCircularPattern4(
                instanceCount, // 参数 1: Number - 实例总数
                (360.0 * Math.PI / 180.0), // 参数 2: Spacing - 阵列总角度 (360度，需转为弧度)
                false, // 参数 3: FlipDirection - 反向
                "", // 参数 4: DName - 特征尺寸名称 (通常为空)
                false, // 参数 5: GeometryPattern - 设为 false，执行特征阵列
                true, // 参数 6: EqualSpacing - 均布
                false // 参数 7: VaryInstance - 实例随形变化
            )
            .AssertNotNull("创建圆周阵列失败！");

        return model;
    }

    /// <summary>
    /// 在当前选定的边线、面、特征或循环上创建等半径圆角。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="radius">圆角的半径。</param>
    /// <param name="createdFeature">创建的圆角特征。</param>
    public static IModelDoc2 CreateConstantRadiusFillet(this IModelDoc2 model, double radius, out Feature createdFeature)
    {
        var featureManager = model.FeatureManager;

        // 1. 创建圆角特征的“定义”对象
        var filletDef = (ISimpleFilletFeatureData2)featureManager.CreateDefinition((int)swFeatureNameID_e.swFmFillet);

        // 2. 初始化为“等半径圆角”类型
        filletDef.Initialize((int)swSimpleFilletType_e.swConstRadiusFillet);

        // 3. 设置核心参数
        filletDef.DefaultRadius = radius;
        filletDef.PropagateToTangentFaces = true; // 自动延伸到所有相切的面，这是最常用的选项

        // 4. 根据定义创建特征
        createdFeature = featureManager.CreateFeature(filletDef).AssertNotNull("创建等半径圆角失败！");

        return model;
    }

    /// <summary>
    /// 在当前选定的边线上创建“角度-距离”倒角。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="distance">倒角的距离。</param>
    /// <param name="angleDegrees">倒角的角度（单位：度）。</param>
    /// <param name="createdFeature">创建的倒角特征。</param>
    public static IModelDoc2 CreateAngleDistanceChamfer(
        this IModelDoc2 model, double distance, double angleDegrees, out Feature createdFeature)
    {
        var featureManager = model.FeatureManager;

        createdFeature = featureManager.InsertFeatureChamfer(
                // Options: 使用 swFeatureChamferOption_e 枚举，启用切线延伸
                Options: (int)swFeatureChamferOption_e.swFeatureChamferTangentPropagation,
                
                // ChamferType: 指定类型为“角度-距离”
                ChamferType: (int)swChamferType_e.swChamferAngleDistance,
                
                // Width: 对于“角度-距离”类型，此参数表示距离
                Width: distance,
                
                // Angle: API 需要弧度
                Angle: angleDegrees * Math.PI / 180.0,
                
                // 其他参数对于“角度-距离”类型不适用，传入0
                OtherDist: 0,
                VertexChamDist1: 0,
                VertexChamDist2: 0,
                VertexChamDist3: 0
            )
            .AssertNotNull("创建角度-距离倒角失败！");

        return model;
    }
}