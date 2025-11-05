// 引入必要的命名空间

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using SolidWorks; // 包含我们的断言扩展方法

// --- 配置区域 ---
const string TargetFilePath = @"D:\User\Desktop\Project\SolidWorksProject\xhzhuji VVFWCRYOAXIS.stp.temp.SLDPRT";
const int ComponentsToKeep = 3000; // 要保留的零部件/实体的数量

// --- 主程序开始 ---
Console.WriteLine("--- SolidWorks 大型装配体/零件简化程序 ---");
var stopwatch = Stopwatch.StartNew();

// 1. 连接到 SolidWorks 实例
Console.WriteLine("正在连接到 SolidWorks 实例...");
var sldWorks = SolidWorksConnector.GetSldWorksApp();
sldWorks.AssertNotNull("未能连接或启动 SolidWorks。");
sldWorks.Visible = true; // 在生产环境中可以注释掉此行以在后台运行，加快速度

// 2. 打开目标文件
Console.WriteLine($"正在打开文件: {Path.GetFileName(TargetFilePath)}");

string extension = Path.GetExtension(TargetFilePath).ToUpperInvariant();
swDocumentTypes_e docTypeToOpen;

// -- 根据文件扩展名确定文档类型 --
if (extension.EndsWith("SLDPRT"))
{
    docTypeToOpen = swDocumentTypes_e.swDocPART;
}
else if (extension.EndsWith("SLDASM"))
{
    docTypeToOpen = swDocumentTypes_e.swDocASSEMBLY;
}
else
{
    // 如果是其他类型，抛出异常，因为我们的简化逻辑不支持
    throw new NotSupportedException($"不支持的文件类型: {extension}。程序仅支持 .SLDPRT 和 .SLDASM 文件。");
}

int errors = 0;
int warnings = 0;
var sourceDoc = sldWorks.OpenDoc6(
    TargetFilePath,
    (int)swDocumentTypes_e.swDocPART, // 我们已知这是个零件
    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
    "",
    ref errors,
    ref warnings) as IModelDoc2;

sourceDoc.AssertNotNull($"无法打开源文件: {TargetFilePath}");
Console.WriteLine("源文件打开成功。");

// 3. 判断文档类型并执行相应的简化逻辑
// 我们的逻辑现在只针对零件，如果需要支持装配体，需要另外编写一个复制零部件到新装配体的方法
if (sourceDoc.GetType() != (int)swDocumentTypes_e.swDocPART)
{
    Console.WriteLine("错误：此版本的脚本仅支持简化零件文件 (.SLDPRT)。");
    return;
}

// 3. 核心简化逻辑：复制到新零件
var newPartDoc = SimplifyPartByCopyToNew(sldWorks, sourceDoc);
newPartDoc.AssertNotNull("创建简化零件失败。");

// 4. 关闭原始文件，不保存
Console.WriteLine($"正在关闭原始文件: {Path.GetFileName(TargetFilePath)}");
sldWorks.CloseDoc(Path.GetFileName(TargetFilePath));

// 5. 保存新的简化模型
string simplifiedFilePath = Path.ChangeExtension(TargetFilePath, $".simplified_top{ComponentsToKeep}_simplest.SLDPRT");
var modelExt = newPartDoc.Extension;
bool saveSuccess = modelExt.SaveAs(simplifiedFilePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
    (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
saveSuccess.AssertTrue($"保存简化模型失败: {simplifiedFilePath}");
Console.WriteLine($"\n简化模型已成功保存至: {simplifiedFilePath}");

// 6. 完成
stopwatch.Stop();
Console.WriteLine($"\n--- 所有任务完成！总耗时: {stopwatch.Elapsed.TotalSeconds:F2} 秒 ---");

return;
// --- 核心逻辑方法 ---

/// <summary>
/// 【最终可靠版】通过将最大的N个实体复制到一个全新的零件文件中来简化模型。
/// 这是处理超大型多实体零件最稳定、最高效的方法。
/// </summary>
static IModelDoc2? SimplifyPartByCopyToNew(ISldWorks sldWorks, IModelDoc2 sourceDoc)
{
    var sourcePartDoc = (IPartDoc)sourceDoc;

    Console.WriteLine("正在从源文件中获取所有实体...");
    object[]? bodies = sourcePartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
    (bodies is not null && bodies.Length > 0).AssertTrue("在源零件中未找到任何实体。");
    Console.WriteLine($"分析开始，共找到 {bodies.Length} 个实体。");

    // 进度条初始化
    int totalBodies = bodies.Length;
    int currentProgress = 0;
    
    // --- 核心变更：计算“重要性评分” ---
    Console.WriteLine("正在计算每个实体的体积和面数，以生成重要性评分...");
    var bodyData = new List<(IBody2 Body, double SignificanceScore)>();
    foreach (var bodyObj in bodies)
    {
        // 更新并显示进度条
        currentProgress++;
        ProgressBar.Write(currentProgress, totalBodies, "计算实体复杂度");
        
        var body = (IBody2)bodyObj;
        
        int faceCount = body.GetFaceCount();
        if (faceCount == 0) continue; // 忽略没有面的实体（如纯线条或点）

        if (body.GetBodyBox() is not double[] box) continue;
        double volume = CalculateVolume(box);
        if (volume <= 0) continue; // 忽略没有体积的实体

        // 计算重要性评分
        double significanceScore = volume / faceCount;
        bodyData.Add((body, significanceScore));
    }

    // --- 核心变更：按重要性评分降序排序，保留最重要的 ---
    Console.WriteLine("正在按重要性评分对实体进行排序...");
    var sortedBodies = bodyData.OrderByDescending(d => d.SignificanceScore).ToList();

    if (sortedBodies.Count <= ComponentsToKeep)
    {
        Console.WriteLine($"实体总数 ({sortedBodies.Count}) 不超过要保留的数量 ({ComponentsToKeep})，无需简化。");
    }

    var bodiesToKeep = sortedBodies
        .Take(ComponentsToKeep)
        .Select(d => d.Body)
        .ToList();

    Console.WriteLine($"排序完成。将把最重要的 {bodiesToKeep.Count} 个实体复制到新零件中。");

    // 步骤 2: 创建一个新的空零件
    Console.WriteLine("正在创建新的目标零件文件...");
    var newPartDoc = (sldWorks.NewPart() as IModelDoc2)
        .AssertNotNull("未能使用 ISldWorks.NewPart() 创建新零件。请检查SolidWorks是否能手动创建新零件。");

    var newPart = (IPartDoc)newPartDoc;

    Console.WriteLine("正在执行实体复制操作...");
    // sourceDoc.Lock(); // 锁定源文档
    // newPartDoc.Lock(); // 锁定目标文档

    // 为复制操作重置进度条
    int totalToCopy = bodiesToKeep.Count;
    currentProgress = 0;
    
    foreach (var bodyToCopy in bodiesToKeep)
    {
        // 更新并显示复制进度
        currentProgress++;
        ProgressBar.Write(currentProgress, totalToCopy, "正在复制实体");
        
        // 步骤 3a: 在内存中创建实体的临时副本
        var tempBody = (IBody2)bodyToCopy.Copy();
        tempBody.AssertNotNull("复制实体到内存失败。");

        // 步骤 3b: 将临时实体副本作为新特征插入到新零件中
        // 第二个参数 `AddToExisting`: 第一个实体设为 false (创建基体)，后续设为 true (作为新实体添加)
        newPart.CreateFeatureFromBody3(tempBody, false, (int)swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify);
    }

    // sourceDoc.UnLock();
    // newPartDoc.UnLock();


    newPartDoc.ForceRebuild3(false); // 重建新零件
    Console.WriteLine("所有实体已成功复制到新零件。");
    return newPartDoc;
}

static double CalculateVolume(double[] box)
{
    var vol = (box[3] - box[0]) * (box[4] - box[1]) * (box[5] - box[2]);
    return double.IsNaN(vol) ? 0 : vol;
}

/// <summary>
/// 一个简单的静态类，用于在控制台绘制文本进度条。
/// </summary>
public static class ProgressBar
{
    private const int BlockCount = 30; // 进度条的宽度（字符数）
    private static readonly object LockObject = new object(); // 用于线程安全

    /// <summary>
    /// 在控制台的同一行上写入或更新进度条。
    /// </summary>
    /// <param name="current">当前进度值</param>
    /// <param name="total">总进度值</param>
    /// <param name="message">显示在进度条前的消息</param>
    public static void Write(int current, int total, string message = "处理中")
    {
        // 使用 lock 确保多线程环境下的控制台输出不会混乱（虽然此应用当前是单线程，但这是个好习惯）
        lock (LockObject)
        {
            // 防止除以零
            if (total == 0) return;

            // 使用 \r (回车符) 将光标移动到行首，实现原地更新的效果
            Console.Write("\r");

            // 计算进度百分比
            double percent = (double)current / total;
            
            // 计算进度条中填充块的数量
            int blocks = (int)(percent * BlockCount);
            
            // 构建进度条字符串
            string progressBar = $"[{new string('=', blocks)}{new string(' ', BlockCount - blocks)}]";

            // 组合最终的输出字符串
            // 使用 PadRight 来确保消息长度一致，防止刷新时留下残影
            string output = $"{message,-20} {progressBar} {percent:P0} ({current}/{total})";

            // 写入控制台
            Console.Write(output);
        }
    }
}

