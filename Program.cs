using SolidWorks.Progress;
using SolidWorks.Utils;

// 连接到 SolidWorks 实例
Console.WriteLine("正在连接到 SolidWorks 实例...");

var sldWorks = SolidWorksConnector.GetSldWorksApp().Setup(sw => sw.CloseSketchAutoGro()).SetVisible(false);

new DrawFlange(sldWorks).Run();