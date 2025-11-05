using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using SolidWorks.Interop.sldworks;

// ReSharper disable InconsistentNaming

// 这个类是为了在现代 .NET 平台上恢复 .NET Framework 中存在的 Marshal.GetActiveObject 功能。
// 由于该功能是 Windows 特有的 COM 操作，并未被包含在跨平台的 .NET 核心库中。
// 因此，我们通过 P/Invoke (平台调用) 直接调用底层的 Windows API (ole32.dll, oleaut32.dll) 来实现相同的功能。
namespace SolidWorks.Utils;

[SupportedOSPlatform("windows")] // 明确指出这个类只在 Windows 上工作
public static class Marshal2
{
    // 定义需要调用的非托管 DLL
    private const string OLE32 = "ole32.dll";
    private const string OLEAUT32 = "oleaut32.dll";

    // 从 Program ID (如 "SldWorks.Application") 获取 CLSID (类标识符)
    [DllImport(OLE32, PreserveSig = false)]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    // 获取一个正在运行的、已在 ROT (Running Object Table) 中注册的 COM 对象
    [DllImport(OLEAUT32, PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);

    /// <summary>
    /// 从给定的 Program ID 获取一个正在运行的 COM 对象的实例。
    /// </summary>
    /// <param name="progID">COM 对象的 Program ID，例如 "Excel.Application" 或 "SldWorks.Application"。</param>
    /// <returns>返回正在运行的 COM 对象实例。</returns>
    public static object GetActiveObject(string progID)
    {
        // 1. 根据 ProgID 获取其唯一的 CLSID
        CLSIDFromProgID(progID, out Guid clsid);

        // 2. 使用 CLSID 从 ROT (Running Object Table) 中获取活动对象
        GetActiveObject(ref clsid, IntPtr.Zero, out object obj);

        return obj;
    }
}

public static class SolidWorksConnector
{
    /// <summary>
    /// 设置 SOLIDWORKS 应用程序的可见性
    /// </summary>
    /// <param name="sldWorks"></param>
    /// <param name="visible"></param>
    /// <returns></returns>
    public static SldWorks SetVisible(this SldWorks sldWorks, bool visible)
    {
        sldWorks.Visible = visible;
        return sldWorks;
    }

    /// <summary>
    /// 获取一个 SOLIDWORKS 应用程序 (SldWorks) 的实例。
    /// 会首先尝试连接到已有的实例，如果失败，则启动一个新实例。
    /// </summary>
    /// <returns>成功则返回 SldWorks 对象，失败则返回 null。</returns>
    public static SldWorks GetSldWorksApp()
    {
        SldWorks? swApp;
        try
        {
            // 首先，尝试连接到一个已经运行的 SOLIDWORKS 实例
            Console.WriteLine("日志：[Connector] 正在尝试连接到已运行的 SOLIDWORKS 实例...");
            swApp = (SldWorks)Marshal2.GetActiveObject("SldWorks.Application");
            Console.WriteLine("日志：[Connector] 成功连接到已运行的实例。");
            return swApp;
        }
        catch (COMException)
        {
            // 如果没有正在运行的实例，则启动一个新的
            Console.WriteLine("日志：[Connector] 未找到正在运行的实例，将启动一个新进程...");
            try
            {
                var swType = Type.GetTypeFromProgID("SldWorks.Application");
                if (swType == null)
                {
                    throw new SolidWorksConnectionException("无法获取 SldWorks.Application 的 ProgID。请确保 SOLIDWORKS 已正确安装。");
                }

                swApp = (SldWorks?)Activator.CreateInstance(swType);
                if (swApp == null)
                {
                    throw new SolidWorksConnectionException("Activator.CreateInstance 未能创建 SOLIDWORKS 实例。");
                }

                swApp.Visible = true;
                Console.WriteLine("日志：[Connector] 新的 SOLIDWORKS 实例已启动并可见。");
                return swApp;
            }
            catch (Exception ex)
            {
                throw new SolidWorksConnectionException($"启动 SOLIDWORKS 时发生严重错误: {ex.Message}", ex);
            }
        }
    }
}

/// <summary>
/// 表示在连接到 SOLIDWORKS 应用程序期间发生的特定错误。
/// </summary>
public class SolidWorksConnectionException : Exception
{
    public SolidWorksConnectionException() { }
    public SolidWorksConnectionException(string message) : base(message) { }
    public SolidWorksConnectionException(string message, Exception inner) : base(message, inner) { }
}