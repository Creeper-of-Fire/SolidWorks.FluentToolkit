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
    /// 将另一个点（或向量）加到当前点上 (this + other)。
    /// 这常用于将一个点沿着一个向量进行平移。
    /// </summary>
    /// <param name="other">要加上的点或向量。</param>
    /// <returns>表示和的新 Point3D 对象。</returns>
    public Point3D Add(Point3D other)
    {
        // 对每个分量执行加法
        return new Point3D(this.X + other.X, this.Y + other.Y, this.Z + other.Z);
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

    /// <summary>
    /// 将当前点视为向量，并按给定的标量进行缩放。
    /// 这会改变向量的长度（模）。
    /// </summary>
    /// <param name="scalar">用于缩放的乘数。</param>
    /// <returns>缩放后的新 Point3D 向量。</returns>
    public Point3D Scale(double scalar)
    {
        // 将每个分量乘以标量
        return new Point3D(this.X * scalar, this.Y * scalar, this.Z * scalar);
    }

    /// <summary>
    /// 将当前点视为向量，并计算其单位向量（长度为1的向量）。
    /// 单位向量只保留方向信息，长度被归一化为1。
    /// </summary>
    /// <returns>与原始向量方向相同但长度为1的新 Point3D 向量。如果原始向量为零向量，则返回零向量。</returns>
    public Point3D Normalize()
    {
        // 计算向量的模长 (Magnitude)
        double magnitude = Math.Sqrt(this.X * this.X + this.Y * this.Y + this.Z * this.Z);

        // 健壮性检查：如果模长非常接近于0，则无法归一化，直接返回零向量以避免除以零的错误。
        if (magnitude < 1e-9)
        {
            return new Point3D(0, 0, 0);
        }

        // 每个分量除以模长，得到单位向量
        return new Point3D(this.X / magnitude, this.Y / magnitude, this.Z / magnitude);
    }
}