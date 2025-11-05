using SolidWorks.Interop.sldworks;

namespace SolidWorks.Progress;

public abstract class AbstractRun(SldWorks sldWorks)
{
    protected SldWorks SldWorks { get; } = sldWorks;
    public virtual void Run() { }
}