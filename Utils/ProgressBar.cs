namespace SolidWorks.Utils;

/// <summary>
/// 一个简单的静态类，用于在控制台绘制文本进度条。
/// </summary>
public static class ProgressBar
{
    private const int BlockCount = 30; // 进度条的宽度（字符数）
    private static readonly Lock LockObject = new(); // 用于线程安全

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