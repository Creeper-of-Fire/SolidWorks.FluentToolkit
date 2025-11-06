using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorks.Utils;

public static class SolidWorksExtensions
{
    /// <summary>
    /// 关闭捕捉到模型几何体，防止吸附，以确保绘图的绝对精度
    /// </summary>
    /// <param name="sldWorks"></param>
    /// <returns></returns>
    public static ISldWorks CloseSketchAutoGro(this ISldWorks sldWorks)
    {
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInferFromModel, false);
        return sldWorks;
    }
}