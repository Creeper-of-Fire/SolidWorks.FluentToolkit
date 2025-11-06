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
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInferFromModel, false);
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchSnapsGrid, true);
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchSnapsPoints, true);
        sldWorks.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchAutomaticRelations, true);
        return sldWorks;
    }
}