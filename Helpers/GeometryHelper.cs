using SolidWorks.Helpers.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Utils;

namespace SolidWorks.Helpers;

public static class GeometryHelper
{
    // 定义一个微小的浮点数容差，用于比较几何参数
    private const double Epsilon = 1e-9;

    /// <summary>
    /// 从一个实体上查找所有面，并返回第一个符合条件的圆柱面。
    /// </summary>
    /// <param name="body">要搜索的实体对象。</param>
    /// <param name="radius">所需圆柱面的半径。</param>
    /// <returns>找到的 IFace2 对象，如果未找到则返回 null。</returns>
    public static IFace2? FindCylindricalFaceByRadius(IBody2 body, double radius)
    {
        var faces = (object[])body.GetFaces();
        foreach (IFace2 face in faces)
        {
            var surface = (ISurface)face.GetSurface();
            if (!surface.IsCylinder()) continue;

            double[]? parameters = (double[])surface.CylinderParams;
            if (Math.Abs(parameters[6] - radius) < Epsilon)
            {
                return face;
            }
        }

        return null;
    }

    /// <summary>
    /// 从一个实体上查找所有面，并返回第一个符合条件的平面。
    /// </summary>
    /// <param name="body">要搜索的实体对象。</param>
    /// <param name="planeNormal">所需平面的法线向量。</param>
    /// <param name="pointOnPlane">所需平面上的一个点。</param>
    /// <returns>找到的 IFace2 对象，如果未找到则返回 null。</returns>
    public static IFace2? FindPlanarFace(IBody2 body, Point3D planeNormal, Point3D pointOnPlane)
    {
        var faces = (object[])body.GetFaces();
        foreach (IFace2 face in faces)
        {
            var surface = (ISurface)face.GetSurface();
            if (!surface.IsPlane()) continue;

            var parameters = (double[])surface.PlaneParams; // 获取平面参数
            var normal = new Point3D(parameters[0], parameters[1], parameters[2]);
            var rootPoint = new Point3D(parameters[3], parameters[4], parameters[5]);

            // 检查法线方向是否一致 (或相反)
            if (Math.Abs(normal.Dot(planeNormal)) > 1.0 - Epsilon)
            {
                // 检查平面是否包含指定的点。
                // 计算从平面原点到目标点的向量，如果该向量与法线垂直，则点在平面上。
                var vectorInPlane = pointOnPlane.Subtract(rootPoint);
                if (Math.Abs(vectorInPlane.Dot(normal)) < Epsilon)
                {
                    return face;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// 查找两个面之间共享的唯一一条边线。
    /// </summary>
    /// <param name="face1">第一个面。</param>
    /// <param name="face2">第二个面。</param>
    /// <returns>共享的 IEdge 对象，如果未找到则返回 null。</returns>
    public static IEdge? GetIntersectionEdge(IFace2 face1, IFace2 face2)
    {
        var edgesOfFace1 = (object[])face1.GetEdges();
        var face2Ptr = (object)face2; // 用于比较指针

        foreach (IEdge edge in edgesOfFace1)
        {
            var adjacentFaces = (object[])edge.GetTwoAdjacentFaces2();
            // 检查 face2 是否是这条边的两个相邻面之一
            if (adjacentFaces.Any(adjFace => adjFace == face2Ptr))
            {
                return edge;
            }
        }

        return null;
    }

    /// <summary>
    /// 获取一个特征创建的第一个实体(Body)。
    /// 对于大多数凸台/基体特征，这非常有效。
    /// </summary>
    public static IBody2 GetFirstBody(this Feature feature)
    {
        var faces = (object[])feature.GetFaces();
        if (faces == null || faces.Length == 0)
        {
            throw new InvalidOperationException($"特征 '{feature.Name}' 不包含任何面，无法获取实体。");
        }

        return (IBody2)((IFace2)faces[0]).GetBody().AssertNotNull("未能从特征的面中获取实体。");
    }

    /// <summary>
    /// 通过其一个圆形边界的属性来查找一个特定的锥形面。
    /// </summary>
    /// <param name="body">要搜索的实体。</param>
    /// <param name="circleRadius">边界圆的半径。</param>
    /// <param name="circleCenter">边界圆的中心点坐标。</param>
    /// <returns>找到的 IFace2 对象，如果未找到则返回 null。</returns>
    public static IFace2? FindConicalFaceByBoundaryCircle(IBody2 body, double circleRadius, Point3D circleCenter)
    {
        var faces = (object[])body.GetFaces();
        foreach (IFace2 face in faces)
        {
            var surface = (ISurface)face.GetSurface();
            if (!surface.IsCone()) continue; // 只关心锥形面

            // 检查该锥形面的所有边界边
            var edges = (object[])face.GetEdges();
            foreach (IEdge edge in edges)
            {
                var curve = (ICurve)edge.GetCurve();
                if (!curve.IsCircle()) continue; // 只关心圆形边界

                var parameters = (double[])curve.CircleParams;
                var center = new Point3D(parameters[0], parameters[1], parameters[2]);
                double radius = parameters[6];

                if (Math.Abs(radius - circleRadius) < Epsilon &&
                    Math.Abs(center.X - circleCenter.X) < Epsilon &&
                    Math.Abs(center.Y - circleCenter.Y) < Epsilon &&
                    Math.Abs(center.Z - circleCenter.Z) < Epsilon)
                {
                    // 找到了完全匹配的锥形面
                    return face;
                }
            }
        }

        return null; // 遍历完所有面都未找到
    }
}