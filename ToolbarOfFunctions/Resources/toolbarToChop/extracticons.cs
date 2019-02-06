using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Interop.Excel;
using stdole;

public class OfficeIconUtils
{
    public static void ExtractAllIcons(string xlsmPath, string targetFolder)
    {
        // extract  customUI.xml
        var zf = new ZipFile(xlsmPath);
        var entry = zf.GetEntry("customUI/customUI.xml");
        var zipStream = zf.GetInputStream(entry);
        XNamespace ns = "http://schemas.microsoft.com/office/2006/01/customui";
        var root = XElement.Load(zipStream);
        foreach (var gallery in root.Descendants(ns + "gallery"))
        {
            //create a sub-folder for the gallery
            var subFolder = Path.Combine(targetFolder,
                gallery.Attribute("label").Value);
            var width = int.Parse(gallery.Attribute("itemWidth").Value);
            var height = int.Parse(gallery.Attribute("itemHeight").Value);
            Directory.CreateDirectory(subFolder);
            foreach (var item in gallery.Descendants(ns + "item"))
            {
                SaveIcon(item.Attribute("imageMso").Value,
                    subFolder, width, height);
            }
        }
    }

    public static void SaveIcon(string msoName, string folder,
        int width = 32, int height = 32)
    {
        ConvertPixelByPixel(
            ((Application)(ExcelDnaUtil.Application))
                .CommandBars.GetImageMso(msoName, width, height))
            .Save(Path.Combine(folder, string.Format("{0}.png",
            msoName)), ImageFormat.Png);
    }


    public static Bitmap ConvertPixelByPixel(IPictureDisp ipd)
    {
        // get the info about the HBITMAP inside the IPictureDisp
        var dibsection = new DIBSECTION();
        GetObjectDIBSection((IntPtr)ipd.Handle, Marshal.SizeOf(dibsection), ref dibsection);
        var width = dibsection.dsBm.bmWidth;
        var height = dibsection.dsBm.bmHeight;

        // create the destination Bitmap object
        var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);

        unsafe
        {
            // get a pointer to the raw bits
            var pBits = (RGBQUAD*)(void*)dibsection.dsBm.bmBits;

            // copy each pixel manually
            for (var x = 0; x < dibsection.dsBmih.biWidth; x++)
                for (var y = 0; y < dibsection.dsBmih.biHeight; y++)
                {
                    var offset = y * dibsection.dsBmih.biWidth + x;
                    if (pBits[offset].rgbReserved != 0)
                    {
                        bitmap.SetPixel(x, y, Color.FromArgb(pBits[offset].rgbReserved, pBits[offset].rgbRed, pBits[offset].rgbGreen, pBits[offset].rgbBlue));
                    }
                }
        }

        return bitmap;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RGBQUAD
    {
        public byte rgbBlue;
        public byte rgbGreen;
        public byte rgbRed;
        public byte rgbReserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BITMAP
    {
        public Int32 bmType;
        public Int32 bmWidth;
        public Int32 bmHeight;
        public Int32 bmWidthBytes;
        public Int16 bmPlanes;
        public Int16 bmBitsPixel;
        public IntPtr bmBits;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BITMAPINFOHEADER
    {
        public int biSize;
        public int biWidth;
        public int biHeight;
        public Int16 biPlanes;
        public Int16 biBitCount;
        public int biCompression;
        public int biSizeImage;
        public int biXPelsPerMeter;
        public int biYPelsPerMeter;
        public int biClrUsed;
        public int bitClrImportant;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DIBSECTION
    {
        public BITMAP dsBm;
        public BITMAPINFOHEADER dsBmih;
        public int dsBitField1;
        public int dsBitField2;
        public int dsBitField3;
        public IntPtr dshSection;
        public int dsOffset;
    }

    [DllImport("gdi32.dll", EntryPoint = "GetObject")]
    public static extern int GetObjectDIBSection(IntPtr hObject, int nCount, ref DIBSECTION lpObject);

}
{code}