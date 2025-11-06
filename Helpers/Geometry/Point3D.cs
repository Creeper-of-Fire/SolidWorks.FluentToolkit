namespace SolidWorks.Helpers.Geometry;


/// <summary>
/// 表示一个三维空间中的点，使用全局坐标系。
/// 使用 record struct 以获得值类型的高效性和不变性。
/// </summary>
public readonly record struct Point3D(double X, double Y, double Z);