using SolidWorks.Interop.swconst;
using SolidWorks.Utils;
using IBody2 = SolidWorks.Interop.sldworks.IBody2;
using IModelDoc2 = SolidWorks.Interop.sldworks.IModelDoc2;
using IPartDoc = SolidWorks.Interop.sldworks.IPartDoc;
using ISldWorks = SolidWorks.Interop.sldworks.ISldWorks;

namespace SolidWorks.Tools;

public static class SimplifyPart
{
    /// <summary>
    /// 通过将最大的N个实体复制到一个全新的零件文件中来简化模型。
    /// 这是处理超大型多实体零件最稳定、最高效的方法。
    /// </summary>
    public static IModelDoc2 SimplifyPartByCopyToNew(ISldWorks sldWorks, IModelDoc2 sourceDoc, int componentsToKeep)
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

        if (sortedBodies.Count <= componentsToKeep)
        {
            Console.WriteLine($"实体总数 ({sortedBodies.Count}) 不超过要保留的数量 ({componentsToKeep})，无需简化。");
        }

        var bodiesToKeep = sortedBodies
            .Take(componentsToKeep)
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
}