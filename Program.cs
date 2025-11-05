using SolidWorks.Interop.swconst;
using SolidWorks.Progress;
using SolidWorks.Utils;

// 连接到 SolidWorks 实例
Console.WriteLine("正在连接到 SolidWorks 实例...");

var app = SolidWorksConnector.GetSldWorksApp().SetVisible(true).ToApplication();

new DrawComplexFlange(app).Run();