using SolidWorks.Interop.swconst;
using SolidWorks.Progress;
using SolidWorks.Utils;

// 连接到 SolidWorks 实例
Console.WriteLine("正在连接到 SolidWorks 实例...");

var app = SolidWorksConnector.GetSldWorksApp().Setup(sw => sw.CloseSketchAutoGro()).SetVisible(true).ToApplication();

new DrawFlange(app).Run();