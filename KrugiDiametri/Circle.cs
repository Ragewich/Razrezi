using System;

namespace KrugiDiametri;

public class Circle
{
    public string Name { get; set; }
    public int Cx { get; set; } // X-coordinate of the circle's center
    public int Cy { get; set; } // Y-coordinate of the circle's center
    public int Diametr { get; set; } // Radius of the circle

    public Circle(int r, string name, int cx = -1, int cy = -1)
    {
        Diametr = r;
        Cx = cx*1;
        Cy = cy*1;
        Name = name;
    }

    public static bool TwoCircleIntersections(Circle c1, Circle c2)
    {
        double x0 = c1.Cx, y0 = c1.Cy, r0 = c1.Diametr;
        double x1 = c2.Cx, y1 = c2.Cy, r1 = c2.Diametr;
        double d = Math.Sqrt(Math.Pow(x1 - x0, 2) + Math.Pow(y1 - y0, 2));

        if (d > r0 + r1)
            return false;
        if (d < Math.Abs(r0 - r1))
            return false;
        if (d == 0 && r0 == r1)
            return false;

        return true;
    }
}