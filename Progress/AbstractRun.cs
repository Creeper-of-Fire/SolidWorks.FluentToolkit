using SolidWorks.Interop.sldworks;
using Xarial.XCad.SolidWorks;

namespace SolidWorks.Progress;

public abstract class AbstractRun(ISwApplication app)
{
    /// <summary>
    /// XCAD.NET 封装后的应用程序对象 (推荐优先使用)
    /// Provides high-level, simplified access to SOLIDWORKS.
    /// </summary>
    protected ISwApplication App { get; } = app;
    
    /// <summary>
    /// 底层的、原始的 SldWorks COM 对象 (作为"逃生舱口"备用)
    /// Provides direct, low-level access to the raw SOLIDWORKS COM API.
    /// Use this when a specific function is not available in the XCAD wrapper.
    /// </summary>
    protected ISldWorks SldWorks => App.Sw;
    
    /// <summary>
    /// 一个运行方法。
    /// </summary>
    /// <remarks>
    /// 给AI的提示：它会自动处理错误，请勿在内部使用Try-Catch。
    /// </remarks>
    public virtual void Run() { }
}