using SolidWorks.Interop.sldworks;
using SolidWorks.Utils;

namespace SolidWorks.Helpers;

/// <summary>
/// 为 IModelDoc2 提供选择内容的扩展方法。
/// </summary>
public static class SelectHelper
{
    /// <summary>
    /// (流式) 通过直接传递对象来选择一个实体 (如特征、面、边线等)。
    /// 这是最可靠的选择方法，因为它不依赖于名称或类型字符串。
    /// </summary>
    public static IModelDoc2 SelectObject(this IModelDoc2 model, object objectToSelect, bool append = true)
    {
        var entity = (IEntity)objectToSelect;

        // 调用对象自身的 Select4 方法，这是最直接和健壮的方式
        entity.Select4(append, null).AssertTrue($"通过对象进行选择失败。对象: {objectToSelect}");

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