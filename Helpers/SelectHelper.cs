using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Helpers;

/// <summary>
/// 为 IModelDoc2 提供选择内容的扩展方法。
/// </summary>
public static class SelectHelper
{
    /// <inheritdoc cref="SelectByRay(IModelDoc2,Point3D,Point3D,double,swSelectType_e,bool,int)"/>
    public static IModelDoc2 SelectByRay(
        this IModelDoc2 model,
        (double x, double y, double z) start,
        (double dirX, double dirY, double dirZ) direction,
        double rayRadius = 0.001d,
        swSelectType_e filter = swSelectType_e.swSelFACES,
        bool append = false,
        int mark = 0)
    {
        return model.SelectByRay(
            new Point3D(start.x, start.y, start.z),
            new Point3D(direction.dirX, direction.dirY, direction.dirZ),
            rayRadius,
            filter, append, mark);
    }

    /// <summary>
    /// 通过发射射线来选择实体。
    /// </summary>
    /// <param name="model">要进行操作的 ModelDoc2 文档。</param>
    /// <param name="start">射线起点的坐标。</param>
    /// <param name="direction">射线的方向向量。</param>
    /// <param name="rayRadius">射线的半径。对于面和边，此值通常被忽略。</param>
    /// <param name="filter">要选择的实体类型，来自 swSelectType_e 枚举。</param>
    /// <param name="append">是否追加到当前选择集。</param>
    /// <param name="mark">选择标记。</param>
    /// <returns>返回传入的 IModelDoc2 对象，以支持链式调用。</returns>
    public static IModelDoc2 SelectByRay(
        this IModelDoc2 model,
        Point3D start,
        Point3D direction,
        double rayRadius = 0.001d, // 对于面和边，此值通常被忽略，设一个较小值即可
        swSelectType_e filter = swSelectType_e.swSelFACES,
        bool append = false,
        int mark = 0)
    {
        var swDocExt = model.Extension;
        bool result = swDocExt.SelectByRay(
            WorldX: start.X,
            WorldY: start.Y,
            WorldZ: start.Z,
            RayVecX: direction.X,
            RayVecY: direction.Y,
            RayVecZ: direction.Z,
            RayRadius: rayRadius, // 对于面和边，此值通常被忽略，设一个较小值即可
            TypeWanted: (int)filter,
            Append: append,
            Mark: mark,
            Option: (int)swSelectOption_e.swSelectOptionDefault
        );

        result.AssertTrue($"使用 SelectByRay 选择实体失败。" +
                          $"起点:({start.X},{start.Y},{start.Z}), " +
                          $"方向:({direction.X},{direction.Y},{direction.Z}), " +
                          $"类型:{filter}");

        return model;
    }

    /// <summary>
    /// (流式) 通过直接传递对象来选择一个实体 (如特征、面、边线等)。
    /// 这是最可靠的选择方法，因为它不依赖于名称或类型字符串。
    /// </summary>
    public static IModelDoc2 SelectObject(this IModelDoc2 model, object objectToSelect, bool append = true, int mark = 0)
    {
        var entity = (IEntity)objectToSelect;
        
        var selectionMgr = (ISelectionMgr)model.SelectionManager;
        
        var selectData = selectionMgr.CreateSelectData();

        // 调用对象自身的 Select4 方法，这是最直接和健壮的方式
        entity.Select4(append, selectData).AssertTrue($"通过对象进行选择失败。对象: {objectToSelect}");

        return model;
    }

    /// <summary>
    /// (流式) 通过名称选择一个对象，例如基准面。
    /// </summary>
    public static IModelDoc2 SelectByName(this IModelDoc2 model, string name, string type, bool append = false)
    {
        model.Extension.SelectByID2(name, type, 0, 0, 0, append, 0, null, 0)
            .AssertTrue($"通过名称 '{name}' 选择类型为 '{type}' 的对象失败。");
        return model;
    }

    /// <summary>
    /// (流式) 通过坐标选择一个面。
    /// </summary>
    public static IModelDoc2 SelectFaceByPoint(this IModelDoc2 model, double x, double y, double z, bool append = false)
    {
        model.Extension.SelectByID2("", "FACE", x, y, z, append, 0, null, 0)
            .AssertTrue($"在坐标 ({x},{y},{z}) 选择面失败。");
        return model;
    }

    /// <summary>
    /// (流式) 通过坐标选择一条边线 (用于阵列轴等)。
    /// </summary>
    public static IModelDoc2 SelectEdgeByPoint(this IModelDoc2 model, double x, double y, double z, bool append = false, int mark = 1)
    {
        model.Extension.SelectByID2("", "EDGE", x, y, z, append, mark, null, 0)
            .AssertTrue($"在坐标 ({x},{y},{z}) 选择边线失败。");
        return model;
    }

    /// <summary>
    /// (流式) 通过名称选择一个特征 (用于阵列源等)。
    /// </summary>
    public static IModelDoc2 SelectFeature(this IModelDoc2 model, Feature featureToSelect, bool append = true, int mark = 4)
    {
        model.Extension.SelectByID2(featureToSelect.Name, "BODYFEATURE", 0, 0, 0, append, mark, null, 0)
            .AssertTrue($"选择要阵列的特征 '{featureToSelect.Name}' 失败！");
        return model;
    }


    /// <summary>
    /// (流式) 清除所有当前选择。这是链中的一个显式步骤。
    /// </summary>
    public static IModelDoc2 ClearSelect(this IModelDoc2 model)
    {
        model.ClearSelection2(true);
        return model;
    }
}