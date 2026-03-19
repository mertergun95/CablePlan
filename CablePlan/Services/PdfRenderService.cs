using System;
using System.IO;
using System.Windows.Media.Imaging;
using PdfiumViewer;

namespace CablePlan.Services;

public static class PdfRenderService
{
    public static BitmapSource RenderPageToBitmapSource(string pdfPath, int pageIndex, int rotationDeg, int targetWidthPx = 2400)
    {
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException(pdfPath);

        using var doc = PdfDocument.Load(pdfPath);
        if (pageIndex < 0 || pageIndex >= doc.PageCount) pageIndex = 0;

        var size = doc.PageSizes[pageIndex];
        double scale = targetWidthPx / size.Width;

        int w = targetWidthPx;
        int h = (int)Math.Round(size.Height * scale);

        var rot = rotationDeg switch
        {
            90 => PdfRotation.Rotate90,
            180 => PdfRotation.Rotate180,
            270 => PdfRotation.Rotate270,
            _ => PdfRotation.Rotate0
        };

        using var img = doc.Render(pageIndex, w, h, 96, 96, rot, PdfRenderFlags.Annotations);
        using var bmp = new System.Drawing.Bitmap(img);
        return BitmapToBitmapSource(bmp);
    }
    private static BitmapSource BitmapToBitmapSource(System.Drawing.Bitmap bitmap)
    {
        using var ms = new MemoryStream();
        bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        ms.Position = 0;

        var bi = new BitmapImage();
        bi.BeginInit();
        bi.CacheOption = BitmapCacheOption.OnLoad;
        bi.StreamSource = ms;
        bi.EndInit();
        bi.Freeze();
        return bi;
    }
}
