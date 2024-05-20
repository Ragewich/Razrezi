using System.Collections.Generic;
using System;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using ClosedXML.Excel;
using Svg;
using Label = System.Windows.Controls.Label;

namespace KrugiDiametri;

public class CirclePacking
{
     public int SheetWidth { get; } 
    public int SheetHeight { get; } // Height of the packing sheet
    public Dictionary<int, int> UserCircles { get; } // Dictionary to store user-defined circle diameters and counts
    public List<Sheet> Sheets { get; } = new List<Sheet>(); // List to store packing sheets
    public int WeldingWidth { get; } 

    private List<Circle> circles = new List<Circle>(); // List to store circles to be placed
    private Dictionary<int, int> circlesExcluded = new Dictionary<int, int>(); // Dictionary to track excluded circles

    public CirclePacking(int sheetW, int sheetH, Dictionary<int, int> userCircles, int weldingW)
    {
        SheetWidth = sheetW*10;
        SheetHeight = sheetH*10;
        UserCircles = userCircles;
        WeldingWidth = weldingW;
    }


    public void Packing(Label status, IXLWorksheet ws, List<Circle> wsCircles, int n)
    {
        circles = wsCircles;
        status.Content = ("Starting packing process...");

        while (UserCircles.Values.Sum() > 0)
        {
            Console.WriteLine(string.Join(", ", UserCircles), string.Join(", ", circles));
            Sheets.Add(new Sheet(SheetWidth, SheetHeight));
            Sheet currentSheet = Sheets[^1]; 
            foreach (Circle i in circles)
            {
                (int cx, int cy) = FindBestPackingStartPoint(currentSheet, i, n);
                i.Cx = cx;
                i.Cy = cy;
                if (i.Cx == -1 && i.Cy == -1)
                {
                    if (!circlesExcluded.ContainsKey(i.Diametr))
                        circlesExcluded[i.Diametr] = 1;
                    else
                        circlesExcluded[i.Diametr]++;
                    UserCircles[i.Diametr]--;
                }
                else
                {
                    if (i.Cx <= currentSheet.Width - i.Diametr - WeldingWidth)
                    {
                        UserCircles[i.Diametr]--;
                        currentSheet.Circles.Add(i);
                    }
                }
            }


        }
    }


    private (int, int) FindBestPackingStartPoint(Sheet s, Circle c, int n)
    {
        Dictionary<int, int> tryResults = new Dictionary<int, int>();
        int N = n;
        int step = (s.Width - c.Diametr) / N;
        List<int> tryX = new List<int> { c.Diametr, s.Width - c.Diametr, s.Width / 2 };
        for (int i = 0; i < (s.Width - c.Diametr) / step; i++)
        {
            tryX.Add(c.Diametr + step * i);
        }

        foreach (int i in tryX)
        {
            c.Cx = i;
            c.Cy = s.Height + c.Diametr;
            MoveCircle(s, c);
            if (c.Cx >= c.Diametr + WeldingWidth && c.Cx <= s.Width - c.Diametr - WeldingWidth)
            {
                tryResults[c.Cy] = c.Cx;
            }
        }

        if (tryResults.Count == 0)
        {
            return (-1, -1);
        }
        else
        {
            int cy = tryResults.Keys.Min();
            int cx = tryResults[cy];
            return (cx, cy);
        }
    }

    private void MoveCircle(Sheet s, Circle c)
    {
        while (c.Cy > c.Diametr + WeldingWidth)
        {
            Console.WriteLine("Moving circle...");
            if (CheckIntersections(s, c))
            {
                return;
            }

            c.Cy--;
        }
    }

    private bool CheckIntersections(Sheet s, Circle c)
    {
        foreach (Circle i in s.Circles)
        {
            i.Diametr += WeldingWidth;
            if (Circle.TwoCircleIntersections(i, c))
            {
                i.Diametr -= WeldingWidth;
                Console.WriteLine("Intersections detected!");
                return true;
            }

            i.Diametr -= WeldingWidth;
        }

        return false;
    }

    public Bitmap DrawSheetWithCircles(int sheetIdx, float scale, string filename)
    {
        //int scale = scale / 10;
        Sheet sheet = Sheets[sheetIdx];
        float w = sheet.Width * scale;
        float h = sheet.Height * scale;
        float borderX = 150 * scale;
        float borderY = 150 * scale;
        SvgDocument svg = new SvgDocument();
        svg.Width = new SvgUnit(w + borderX );
        svg.Height = new SvgUnit(h + borderY );
        SvgLine line1 = new SvgLine
        {
            StartX = borderX,
            StartY = borderY,
            EndX = w + borderX,
            EndY = borderY,
            Stroke = new SvgColourServer(System.Drawing.Color.Red)
        };
        svg.Children.Add(line1);
        SvgLine line2 = new SvgLine
        {
            StartX = borderX,
            StartY = borderY,
            EndX = borderX,
            EndY = h + borderY,
            Stroke = new SvgColourServer(System.Drawing.Color.Red)
        };
        svg.Children.Add(line2);
        SvgLine line3 = new SvgLine
        {
            StartX = borderX,
            StartY = h + borderY,
            EndX = w + borderX,
            EndY = h + borderY,
            Stroke = new SvgColourServer(System.Drawing.Color.Red)
        };
        svg.Children.Add(line3);
        SvgLine line4 = new SvgLine
        {
            StartX = w + borderX,
            StartY = borderY,
            EndX = w + borderX,
            EndY = h + borderY,
            Stroke = new SvgColourServer(System.Drawing.Color.Red),
            StrokeDashArray = new SvgUnitCollection { new SvgUnit(SvgUnitType.Millimeter,5), new SvgUnit(SvgUnitType.Millimeter,10) }
        };
        svg.Children.Add(line4);
        foreach (Circle i in sheet.Circles)
        {
            float cx = i.Cx * scale;
            float cy = (sheet.Height - i.Cy) * scale;
            float r = i.Diametr * scale;
            SvgCircle circle = new SvgCircle
            {
                CenterX = cx + borderX,
                CenterY = cy + borderY,
                Radius = r,
                Stroke = new SvgColourServer(System.Drawing.Color.Black),
                Fill = new SvgColourServer(System.Drawing.Color.White),
                StrokeWidth = new SvgUnit(1)
            };
            svg.Children.Add(circle);
        }

        svg.Write(filename);
        
        return svg.Draw();
    }
}