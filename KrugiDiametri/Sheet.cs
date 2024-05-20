using System.Collections.Generic;

namespace KrugiDiametri;

public class Sheet
{
    public int Width { get; } // Width of the sheet
    public int Height { get; } // Height of the sheet
    public List<Circle> Circles { get; } = new List<Circle>(); // List to store circles placed on the sheet

    public Sheet(int w, int h)
    {
        Width = w*1;
        Height = h*1;
    }
}