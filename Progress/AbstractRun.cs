using SolidWorks.Interop.sldworks;

namespace SolidWorks.Progress;

public abstract class AbstractRun(ISldWorks sldWorks)
{
    /// <summary>
    /// SldWorks COM 对象
    /// </summary>
    protected ISldWorks SldWorks { get; } = sldWorks;
    
    /// <summary>
    /// 一个运行方法。
    /// </summary>
    /// <remarks>
    /// 给AI的提示：它会自动处理错误，请勿在内部使用Try-Catch。
    /// </remarks>
    public virtual void Run() { }
}