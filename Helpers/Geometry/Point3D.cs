namespace SolidWorks.Helpers.Geometry;

/// <summary>
/// 表示一个三维空间中的点，使用全局坐标系。
/// 使用 record struct 以获得值类型的高效性和不变性。
/// </summary>
public readonly record struct Point3D(double X, double Y, double Z)
{
    /// <summary>
    /// 计算当前点与另一个点之间的向量 (this - other)。
    /// 当 P1.Subtract(P2) 时，结果是从 P2 指向 P1 的向量。
    /// </summary>
    /// <param name="other">要减去的点。</param>
    /// <returns>表示两个点之间差异的新 Point3D 向量。</returns>
    public Point3D Subtract(Point3D other)
    {
        // 对每个分量执行减法
        return new Point3D(this.X - other.X, this.Y - other.Y, this.Z - other.Z);
    }

    /// <summary>
    /// 将当前点视为从原点出发的向量，并计算与另一个向量的点积（内积）。
    /// 点积的结果是一个标量，可以用来判断两个向量的角度关系（例如，结果为0表示垂直）。
    /// </summary>
    /// <param name="other">另一个作为向量的点。</param>
    /// <returns>两个向量的点积（一个标量）。</returns>
    public double Dot(Point3D other)
    {
        // 计算对应分量的乘积之和
        return this.X * other.X + this.Y * other.Y + this.Z * other.Z;
    }
}