using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace SolidWorks;

/// <summary>
/// 包含通用扩展方法的静态类
/// </summary>
public static class AssertionExtensions
{
    /// <summary>
    /// 对布尔值进行断言，确保其为 true。
    /// 如果值为 false，则记录一条详细的错误信息并抛出 InvalidOperationException。
    /// </summary>
    /// <param name="value">要检查的布尔值。</param>
    /// <param name="errorMessage">断言失败时显示的错误信息。</param>
    /// <param name="memberName">【自动捕获】调用此方法的成员名称。</param>
    /// <param name="sourceFilePath">【自动捕获】调用此方法的源文件路径。</param>
    /// <param name="sourceLineNumber">【自动捕获】调用此方法的源文件行号。</param>
    public static void AssertTrue(
        [DoesNotReturnIf(false)] this bool value,
        string errorMessage,
        [CallerMemberName] string memberName = "",
        [CallerFilePath] string sourceFilePath = "",
        [CallerLineNumber] int sourceLineNumber = 0)
    {
        // 如果断言为真 (value is true)，则什么也不做，直接返回。
        if (value)
            return;

        // --- 断言失败的处理逻辑 ---

        // 1. 构建详细的、对调试友好的错误日志
        string fullErrorMessage =
            $"[断言失败] {errorMessage}\n" +
            $"    -> 在方法: {memberName}\n" +
            $"    -> 在文件: {Path.GetFileName(sourceFilePath)}\n" +
            $"    -> 在行号: {sourceLineNumber}";

        // 2. 使用醒目的颜色在控制台输出日志
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(fullErrorMessage);
        Console.ResetColor();

        // 3. 抛出异常，中断程序执行。
        //    使用 InvalidOperationException 是一个不错的选择，因为它表示在当前状态下操作无效。
        throw new InvalidOperationException(fullErrorMessage);
    }

    /// <summary>
    /// 对一个对象进行断言，确保其不为 null。
    /// 如果对象为 null，则记录错误并抛出异常。
    /// 此方法使用 [NotNull] 标注，以通知静态分析器在方法成功返回后，对象是非空的。
    /// </summary>
    /// <typeparam name="T">对象的类型。</typeparam>
    /// <param name="obj">要检查的对象。</param>
    /// <param name="errorMessage">断言失败时显示的错误信息。</param>
    /// <param name="memberName">【自动捕获】调用此方法的成员名称。</param>
    /// <param name="sourceFilePath">【自动捕获】调用此方法的源文件路径。</param>
    /// <param name="sourceLineNumber">【自动捕获】调用此方法的源文件行号。</param>
    /// <returns>返回原始的、非空的对象，以便进行链式调用。</returns>
    [return: NotNull]
    public static T AssertNotNull<T>(
        [NotNull] this T? obj,
        string errorMessage,
        [CallerMemberName] string memberName = "",
        [CallerFilePath] string sourceFilePath = "",
        [CallerLineNumber] int sourceLineNumber = 0) where T : class // 约束为引用类型
    {
        // 如果不为 null，则返回原始对象
        if (obj is not null)
            return obj;

        string fullErrorMessage =
            $"[断言失败] 对象为 Null: {errorMessage}\n" +
            $"    -> 在方法: {memberName}\n" +
            $"    -> 在文件: {Path.GetFileName(sourceFilePath)}\n" +
            $"    -> 在行号: {sourceLineNumber}";

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(fullErrorMessage);
        Console.ResetColor();

        throw new InvalidOperationException(fullErrorMessage);
    }
}