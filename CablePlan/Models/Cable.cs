using System.Collections.Generic;

namespace CablePlan.Models;

public sealed class Cable
{
    public string Id { get; set; } = "";
    public List<PointD> Points { get; set; } = new();
    public PointD LabelPoint { get; set; }
}

public readonly struct PointD
{
    public double X { get; init; }
    public double Y { get; init; }
    public PointD(double x, double y) { X = x; Y = y; }
}


public sealed class PlanData
{
    public string PdfPath { get; set; } = "";
    public string PdfHash { get; set; } = "";
    public int PageIndex { get; set; } = 0;
    public int RotationDeg { get; set; } = 0; // 0/90/180/270
    public List<Cable> Cables { get; set; } = new();
}
