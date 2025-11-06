namespace SolidWorks.Helpers.Geometry;

/// <summary>
/// 一个泛型包装器，用于延迟并重复执行一个委托来获取一个“新鲜”的对象引用。
/// 这在 SolidWorks API 编程中尤其有用，因为在特征操作后，旧的几何对象引用可能会失效。
/// 每次访问 .Value 属性时，都会重新调用工厂委托。
/// </summary>
/// <param name="valueFactory">一个无参数的委托，它知道如何查找并返回所需的对象。</param>
/// <typeparam name="T">要包装的对象类型，例如 IFace2 或 IEdge。</typeparam>
public class LazyRef<T>(Func<T> valueFactory)
    where T : class
{
    private Func<T> ValueFactory { get; } = valueFactory ?? throw new ArgumentNullException(nameof(valueFactory));


    /// <summary>
    /// 获取对象的“新鲜”引用。
    /// 每当访问此属性时，都会重新执行在构造函数中提供的委托。
    /// </summary>
    public T Value => this.ValueFactory();
}