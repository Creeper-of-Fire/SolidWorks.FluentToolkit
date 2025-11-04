// 引入必要的命名空间

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using SolidWorks; // 包含我们的断言扩展方法

// --- 配置区域 ---
const string TargetFilePath = @"D:\User\Desktop\Project\SolidWorksProject\xhzhuji VVFWCRYOAXIS.stp.temp.SLDPRT";
const int ComponentsToKeep = 400; // 要保留的零部件/实体的数量

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
string simplifiedFilePath = Path.ChangeExtension(TargetFilePath, $".simplified_top{ComponentsToKeep}_final.SLDPRT");
var modelExt = newPartDoc.Extension;
bool saveSuccess = modelExt.SaveAs(simplifiedFilePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
    (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
saveSuccess.AssertTrue($"保存简化模型失败: {simplifiedFilePath}");
Console.WriteLine($"\n简化模型已成功保存至: {simplifiedFilePath}");

// 6. 完成
stopwatch.Stop();
Console.WriteLine($"\n--- 所有任务完成！总耗时: {stopwatch.Elapsed.TotalSeconds:F2} 秒 ---");


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

    Console.WriteLine("正在计算每个实体的尺寸...");
    var bodyData = new List<(IBody2 Body, double Volume)>();
    foreach (var bodyObj in bodies)
    {
        var body = (IBody2)bodyObj;
        if (body.GetBodyBox() is not double[] box) continue;
        double volume = CalculateVolume(box);
        if (volume > 0)
        {
            bodyData.Add((body, volume));
        }
    }

    Console.WriteLine("正在对实体进行排序...");
    var sortedBodies = bodyData.OrderByDescending(d => d.Volume).ToList();

    if (sortedBodies.Count <= ComponentsToKeep)
    {
        Console.WriteLine($"实体总数 ({sortedBodies.Count}) 不超过要保留的数量 ({ComponentsToKeep})，无需简化。");
        // 在这种情况下，我们可能还是希望得到一个“干净”的副本，所以继续执行复制流程
    }

    var bodiesToKeep = sortedBodies
        .Take(ComponentsToKeep)
        .Select(d => d.Body)
        .ToList();

    Console.WriteLine($"排序完成。将把最大的 {bodiesToKeep.Count} 个实体复制到新零件中。");

    // 步骤 2: 创建一个新的空零件
    Console.WriteLine("正在创建新的目标零件文件...");
    var newPartDoc = (sldWorks.NewPart() as IModelDoc2)
        .AssertNotNull("未能使用 ISldWorks.NewPart() 创建新零件。请检查SolidWorks是否能手动创建新零件。");

    var newPart = (IPartDoc)newPartDoc;

    Console.WriteLine("正在执行实体复制操作...");
    // sourceDoc.Lock(); // 锁定源文档
    // newPartDoc.Lock(); // 锁定目标文档

    foreach (var bodyToCopy in bodiesToKeep)
    {
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

/// <summary>
/// 根据边界框的6个 double 值计算体积。
/// </summary>
static double CalculateVolume(double[] box)
{
    // box = [Xmin, Ymin, Zmin, Xmax, Ymax, Zmax]
    return (box[3] - box[0]) * (box[4] - box[1]) * (box[5] - box[2]);
}