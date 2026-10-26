using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PaperSwitch.Controls;

/// <summary>內建抓取手勢游標，避免依賴個人游標檔案或外部套件。</summary>
internal static class PreviewCursors
{
    public static Cursor CreateGrab()
    {
        const int size = 32;
        var visual = new DrawingVisual();
        using (var drawing = visual.RenderOpen())
        {
            var hand = Geometry.Parse("M8,21 L5,17 Q3,13 6,12 L9,14 L9,9 Q9,6 12,8 Q13,4 16,7 Q19,5 20,9 Q24,7 24,12 L24,19 Q24,23 21,25 L21,28 L11,28 L11,25 Z");
            drawing.DrawGeometry(Brushes.White, new Pen(Brushes.Black, 1.5) { LineJoin = PenLineJoin.Round }, hand);
            drawing.DrawLine(new Pen(Brushes.Black, 1), new Point(13, 10), new Point(13, 15));
            drawing.DrawLine(new Pen(Brushes.Black, 1), new Point(17, 9), new Point(17, 15));
            drawing.DrawLine(new Pen(Brushes.Black, 1), new Point(21, 11), new Point(21, 16));
        }
        var image = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
        image.Render(visual);
        var pixels = new byte[size * size * 4];
        image.CopyPixels(pixels, size * 4, 0);
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, true);
        // CUR 目錄、32 位元 DIB 與透明遮罩；熱點置於掌心。
        writer.Write((ushort)0); writer.Write((ushort)2); writer.Write((ushort)1);
        writer.Write((byte)size); writer.Write((byte)size); writer.Write((ushort)0);
        writer.Write((ushort)15); writer.Write((ushort)15);
        writer.Write(40 + pixels.Length + size * 4); writer.Write(22);
        writer.Write(40); writer.Write(size); writer.Write(size * 2);
        writer.Write((ushort)1); writer.Write((ushort)32); writer.Write(0); writer.Write(pixels.Length);
        writer.Write(0); writer.Write(0); writer.Write(0); writer.Write(0);
        for (var row = size - 1; row >= 0; row--) writer.Write(pixels, row * size * 4, size * 4);
        for (var row = size - 1; row >= 0; row--)
        {
            for (var block = 0; block < 4; block++)
            {
                byte mask = 0;
                for (var bit = 0; bit < 8; bit++)
                    if (pixels[(row * size + block * 8 + bit) * 4 + 3] == 0) mask |= (byte)(0x80 >> bit);
                writer.Write(mask);
            }
        }
        stream.Position = 0;
        return new Cursor(stream);
    }
}
