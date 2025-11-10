using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Tools;
using SolidWorks.Utils;

namespace SolidWorks.Progress;

public class SimplifyPartByComplex(
    ISldWorks sldWorks,
    string targetFilePath = SimplifyPartByComplex.DefaultTargetFilePath,
    int componentsToKeep = SimplifyPartByComplex.DefaultComponentsToKeep
) : AbstractRun(sldWorks)
{
    private const string DefaultTargetFilePath = @"D:\User\Desktop\Project\SolidWorksProject\xhzhuji VVFWCRYOAXIS.stp.temp.SLDPRT";
    private const int DefaultComponentsToKeep = 3000;

    private string TargetFilePath { get; } = targetFilePath;
    private int ComponentsToKeep { get; } = componentsToKeep;

    public override void Run()
    {
        // --- 主程序开始 ---
        Console.WriteLine("--- SolidWorks 大型装配体/零件简化程序 ---");
        var stopwatch = Stopwatch.StartNew();

        // 2. 打开目标文件
        Console.WriteLine($"正在打开文件: {Path.GetFileName(this.TargetFilePath)}");

        int errors = 0;
        int warnings = 0;
        var sourceDoc = this.SldWorks.OpenDoc6(this.TargetFilePath,
            (int)swDocumentTypes_e.swDocPART, // 我们已知这是个零件
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
            "",
            ref errors,
            ref warnings) as IModelDoc2;

        sourceDoc.AssertNotNull($"无法打开源文件: {this.TargetFilePath}");
        Console.WriteLine("源文件打开成功。");

        // 3. 判断文档类型并执行相应的简化逻辑
        // 我们的逻辑现在只针对零件，如果需要支持装配体，需要另外编写一个复制零部件到新装配体的方法
        if (sourceDoc.GetType() != (int)swDocumentTypes_e.swDocPART)
        {
            Console.WriteLine("错误：此版本的脚本仅支持简化零件文件 (.SLDPRT)。");
            return;
        }

        // 3. 核心简化逻辑：复制到新零件
        var newPartDoc = SimplifyPart.SimplifyPartByCopyToNew(this.SldWorks, sourceDoc, this.ComponentsToKeep);
        newPartDoc.AssertNotNull("创建简化零件失败。");

        // 4. 关闭原始文件，不保存
        Console.WriteLine($"正在关闭原始文件: {Path.GetFileName(this.TargetFilePath)}");
        this.SldWorks.CloseDoc(Path.GetFileName(this.TargetFilePath));

        // 5. 保存新的简化模型
        string simplifiedFilePath = Path.ChangeExtension(this.TargetFilePath, $".simplified_top{this.ComponentsToKeep}_simplest.SLDPRT");
        var modelExt = newPartDoc.Extension;
        bool saveSuccess = modelExt.SaveAs(simplifiedFilePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
            (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
        saveSuccess.AssertTrue($"保存简化模型失败: {simplifiedFilePath}");
        Console.WriteLine($"\n简化模型已成功保存至: {simplifiedFilePath}");

        // 6. 完成
        stopwatch.Stop();
        Console.WriteLine($"\n--- 所有任务完成！总耗时: {stopwatch.Elapsed.TotalSeconds:F2} 秒 ---");
    }
}