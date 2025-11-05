using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorks.Utils;

public static class SolidWorksExtensions
{
    /// <summary>
    /// 关闭自动几何关系，以确保绘图的绝对精度
    /// </summary>
    /// <param name="sldWorks"></param>
    /// <returns></returns>
    public static ISldWorks CloseSketchAutoGro(this ISldWorks sldWorks)
    {
        // sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchAutomaticRelations, false);
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchSnapsGrid, false);
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchSnapsPoints, false);
        return sldWorks;
    }
}