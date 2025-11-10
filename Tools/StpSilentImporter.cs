using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Utils;

namespace SolidWorks.Tools;

/// <summary>
/// 在后台静默地将大型STP文件导入为结构化的SolidWorks装配体。
/// </summary>
public class StpSilentImporter(ISldWorks sldWorks, string stpFilePath)
{
    private readonly ISldWorks _sldWorks = sldWorks;
    private readonly string _stpFilePath = stpFilePath;

    /// <summary>
    /// 执行完整的静默导入流程。
    /// </summary>
    public async Task RunAsync()
    {
        Console.WriteLine("--- SolidWorks STP 静默导入程序 ---");
        var stopwatch = Stopwatch.StartNew();

        // 步骤 1: 前置检查
        File.Exists(_stpFilePath).AssertTrue($"源文件不存在: {_stpFilePath}");
        
        // --- 步骤 1.5: 预验证输出目录 ---
        string outputAsmPath = Path.ChangeExtension(_stpFilePath, ".SLDASM");
        ValidateAndPrepareOutputDirectory(outputAsmPath);

        // 步骤 2: 准备后台任务并显示活动指示器
        Console.WriteLine($"准备导入文件: {Path.GetFileName(_stpFilePath)}");
        Console.WriteLine("导入过程可能需要很长时间，请保持耐心。程序正在后台处理...");

        // 创建一个任务来执行阻塞式的 LoadFile4 API 调用
        var importTask = Task.Run(this.ImportStpInBackground);
        
        // 同时，在主线程中运行一个活动指示器，让用户知道程序没有死锁
        await ShowActivityIndicatorAsync(importTask);

        // 获取导入结果
        var newAssemblyDoc = await importTask;
        newAssemblyDoc.AssertNotNull("LoadFile4 返回了 null，导入失败。请检查SolidWorks系统选项中的导入设置。");

        Console.WriteLine("\n文件已成功加载到内存中。");

        // 步骤 3: 保存新生成的装配体和所有零件
        // LoadFile4 只是在内存中创建了文档，我们需要将它们保存到磁盘
        Console.WriteLine($"准备将装配体及所有零件保存至目录: {outputAsmPath}");
        
        await SaveOrRetryInfinitelyAsync(newAssemblyDoc, outputAsmPath);

        Console.WriteLine($"\n装配体和所有关联零件已成功保存。");

        // 步骤 3: 清理
        _sldWorks.CloseDoc(Path.GetFileName(outputAsmPath));
        Console.WriteLine("已关闭新生成的装配体。");

        stopwatch.Stop();
        Console.WriteLine($"\n--- 所有任务完成！总耗时: {stopwatch.Elapsed:g} ---");
    }
    
    /// <summary>
    /// 无限次尝试保存文档，直到成功为止。
    /// </summary>
    private async Task SaveOrRetryInfinitelyAsync(IModelDoc2 docToSave, string filePath)
    {
        int errors = 0;
        int warnings = 0;
        int attemptCount = 1;

        while (true)
        {
            Console.WriteLine($"正在尝试保存... (第 {attemptCount} 次)");
            
            bool success = docToSave.Extension.SaveAs(
                filePath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, ref errors, ref warnings);

            if (success)
            {
                if(warnings != 0)
                {
                    Console.WriteLine($"保存成功，但出现警告，代码: {(swFileLoadWarning_e)warnings}");
                }
                // 成功，跳出无限循环
                break; 
            }
            
            // 如果保存失败，打印错误信息并准备下一次重试
            Console.WriteLine($"第 {attemptCount} 次保存失败。错误代码: {(swFileLoadError_e)errors}, 警告: {(swFileLoadWarning_e)warnings}");
            attemptCount++;
            
            Console.WriteLine("等待 10 秒后重试...");
            await Task.Delay(10000);
        }
    }
    
    /// <summary>
    /// 在开始前验证输出路径，确保目录存在且拥有写入权限。
    /// 这是“快速失败”原则的关键实现。
    /// </summary>
    private void ValidateAndPrepareOutputDirectory(string fullOutputPath)
    {
        Console.WriteLine("正在验证输出目录...");
        
        string? outputDir = Path.GetDirectoryName(fullOutputPath);
        outputDir.AssertNotNull($"无法从输出路径 '{fullOutputPath}' 中提取有效的目录信息。");

        // 1. 确保目录存在（如果不存在则尝试创建）
        try
        {
            Directory.CreateDirectory(outputDir);
        }
        catch (Exception ex)
        {
            throw new IOException($"无法创建输出目录 '{outputDir}'。请检查路径是否合法或是否存在权限问题。", ex);
        }

        // 2. 通过尝试写入一个临时文件来验证写入权限
        string tempFile = Path.Combine(outputDir, Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            File.WriteAllText(tempFile, "write_permission_check");
            File.Delete(tempFile);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new UnauthorizedAccessException($"对目标目录 '{outputDir}' 没有写入权限。请以管理员身份运行程序或检查文件夹的安全设置。", ex);
        }
        catch (Exception ex)
        {
            throw new IOException($"测试对目录 '{outputDir}' 的写入权限时发生未知IO错误。", ex);
        }
        
        Console.WriteLine($"输出目录验证成功: {outputDir}");
    }

    /// <summary>
    /// 在后台线程中调用阻塞的 LoadFile4 API。
    /// </summary>
    private IModelDoc2 ImportStpInBackground()
    {
        int errors = 0;
        
        // 这是执行导入的核心API。
        // ArgString: 对于3D Interconnect模式，此参数为空。
        // ImportData: 传入null，让SolidWorks使用在GUI中预设的全局导入选项。
        var modelDoc = _sldWorks.LoadFile4(
            _stpFilePath,
            "", 
            null, 
            ref errors);
            
        // 检查是否有加载错误
        if (errors != 0)
        {
            // 抛出异常而不是返回null，符合“快速失败”原则
            throw new Exception($"LoadFile4 API 调用失败，错误代码: {(swFileLoadError_e)errors}");
        }

        return modelDoc as IModelDoc2;
    }
    
    /// <summary>
    /// 在后台线程中调用阻塞的 SaveAs API。
    /// </summary>
    private bool SaveAssemblyInBackground(IModelDoc2 modelDoc, string savePath)
    {
        int errors = 0;
        int warnings = 0;
        return modelDoc.Extension.SaveAs(
            savePath,
            (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
            null, ref errors, ref warnings);
    }

    /// <summary>
    /// 通过轮询监控输出目录中的文件数量来显示保存进度。
    /// </summary>
    private async Task ShowSaveProgressAsync(Task saveTask, string directoryToMonitor)
    {
        Console.WriteLine("开始保存操作，正在监控文件创建进度...");
        var timer = Stopwatch.StartNew();

        while (!saveTask.IsCompleted)
        {
            int fileCount = 0;
            try
            {
                // 计算目录下已创建的 SolidWorks 零件文件数量
                fileCount = Directory.GetFiles(directoryToMonitor, "*.SLDPRT", SearchOption.TopDirectoryOnly).Length;
            }
            catch (DirectoryNotFoundException)
            {
                // 目录可能尚未创建，这很正常
            }
            
            // 刷新进度显示
            ProgressBar.Write(fileCount, 40000, $"正在保存文件 (已发现 {fileCount} 个)");

            // 等待一秒或直到任务完成
            await Task.Delay(1000);
        }
        
        // 最终清理并显示最终计数
        int finalFileCount = Directory.GetFiles(directoryToMonitor, "*.SLDPRT", SearchOption.TopDirectoryOnly).Length;
        ProgressBar.Write(finalFileCount, finalFileCount, "保存完成"); 
    }


    /// <summary>
    /// 在控制台显示一个简单的活动指示器，直到指定的任务完成。
    /// </summary>
    private async Task ShowActivityIndicatorAsync(Task importTask)
    {
        var spinner = new[] { '/', '-', '\\', '|' };
        int spinnerIndex = 0;
        var timer = Stopwatch.StartNew();

        while (!importTask.IsCompleted)
        {
            // \r 将光标移回行首，实现原地刷新效果
            Console.Write($"\r正在处理中... {spinner[spinnerIndex]}  (已用时: {timer.Elapsed:g})");
            spinnerIndex = (spinnerIndex + 1) % spinner.Length;
            
            // 等待100毫秒或直到任务完成
            await Task.Delay(100);
        }
        
        // 清理最后一行输出
        Console.Write($"\r{new string(' ', Console.WindowWidth - 1)}\r");
        timer.Stop();
    }
}