using SolidWorks.Progress;
using SolidWorks.Utils;

// 1. 连接到 SolidWorks 实例
Console.WriteLine("正在连接到 SolidWorks 实例...");
var sldWorks = SolidWorksConnector.GetSldWorksApp().SetVisible(false);

new SimplifyPartByComplex(sldWorks).Run();