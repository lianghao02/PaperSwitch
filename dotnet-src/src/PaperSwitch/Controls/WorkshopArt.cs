using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PaperSwitch.Controls;

/// <summary>工坊圖集共用快取；素材內嵌於組件，不依賴本機路徑。</summary>
public static class WorkshopArt
{
    public static readonly DependencyProperty IconProperty = DependencyProperty.RegisterAttached(
        "Icon", typeof(int), typeof(WorkshopArt), new FrameworkPropertyMetadata(-1));
    public static int GetIcon(DependencyObject target) => (int)target.GetValue(IconProperty);
    public static void SetIcon(DependencyObject target, int value) => target.SetValue(IconProperty, value);

    public static readonly DependencyProperty MaterialProperty = DependencyProperty.RegisterAttached(
        "Material", typeof(WorkshopMaterial), typeof(WorkshopArt), new FrameworkPropertyMetadata(WorkshopMaterial.Leather));
    public static WorkshopMaterial GetMaterial(DependencyObject target) => (WorkshopMaterial)target.GetValue(MaterialProperty);
    public static void SetMaterial(DependencyObject target, WorkshopMaterial value) => target.SetValue(MaterialProperty, value);

    private static readonly Lazy<BitmapSource> Materials = new(() => Load("material-atlas.png"));
    private static readonly Lazy<BitmapSource> Icons = new(() => Load("icon-atlas.png"));
    private static readonly Dictionary<WorkshopMaterial, BitmapSource[]> SurfaceCache = new();
    private static readonly Dictionary<int, BitmapSource> IconCache = new();

    private static BitmapSource Load(string name)
    {
        var bitmap = new BitmapImage();
        bitmap.BeginInit();
        bitmap.UriSource = new Uri($"pack://application:,,,/PaperSwitch;component/Assets/Workshop/{name}", UriKind.Absolute);
        bitmap.CacheOption = BitmapCacheOption.OnLoad;
        bitmap.EndInit();
        bitmap.Freeze();
        return bitmap;
    }

    private static BitmapSource Crop(BitmapSource source, Int32Rect bounds)
    {
        var bitmap = new CroppedBitmap(source, bounds);
        bitmap.Freeze();
        return bitmap;
    }

    internal static BitmapSource[] GetSurface(WorkshopMaterial material)
    {
        if (SurfaceCache.TryGetValue(material, out var cached)) return cached;
        var source = Materials.Value;
        int tileWidth = source.PixelWidth / 2, tileHeight = source.PixelHeight / 2;
        int tile = (int)material;
        int originX = tile % 2 * tileWidth, originY = tile / 2 * tileHeight;
        // 保留角落的花葉、縫線與包邊，只延伸中央材質。
        int insetX = (int)(tileWidth * 0.22), insetY = (int)(tileHeight * 0.22);
        int[] xs = { 0, insetX, tileWidth - insetX, tileWidth };
        int[] ys = { 0, insetY, tileHeight - insetY, tileHeight };
        var slices = new BitmapSource[9];
        for (int row = 0; row < 3; row++)
            for (int column = 0; column < 3; column++)
                slices[row * 3 + column] = Crop(source, new Int32Rect(
                    originX + xs[column], originY + ys[row], xs[column + 1] - xs[column], ys[row + 1] - ys[row]));
        SurfaceCache[material] = slices;
        return slices;
    }

    internal static BitmapSource? GetIconSource(int index)
    {
        if (index is < 0 or > 15) return null;
        if (IconCache.TryGetValue(index, out var cached)) return cached;
        var source = Icons.Value;
        int col = index % 4;
        int row = index / 4;
        const int inset = 5; // 向內微縮 5px，徹底排除邊界橫線與相鄰圖示溢出
        int x1 = (int)Math.Round((col * source.PixelWidth) / 4.0) + inset;
        int x2 = (int)Math.Round(((col + 1) * source.PixelWidth) / 4.0) - inset;
        int y1 = (int)Math.Round((row * source.PixelHeight) / 4.0) + inset;
        int y2 = (int)Math.Round(((row + 1) * source.PixelHeight) / 4.0) - inset;
        var icon = Crop(source, new Int32Rect(x1, y1, Math.Max(1, x2 - x1), Math.Max(1, y2 - y1)));
        IconCache[index] = icon;
        return icon;
    }
}

public enum WorkshopMaterial { Parchment, Leather, Honey, Moss }

/// <summary>沿用 Border 的版面與子元件行為，以九宮格繪製可縮放的手繪材質。</summary>
public sealed class WorkshopSurface : Border
{
    public static readonly DependencyProperty MaterialProperty = DependencyProperty.Register(
        nameof(Material), typeof(WorkshopMaterial), typeof(WorkshopSurface),
        new FrameworkPropertyMetadata(WorkshopMaterial.Parchment, FrameworkPropertyMetadataOptions.AffectsRender),
        value => value is WorkshopMaterial material && Enum.IsDefined(material));
    public WorkshopMaterial Material { get => (WorkshopMaterial)GetValue(MaterialProperty); set => SetValue(MaterialProperty, value); }

    public static readonly DependencyProperty EdgeSizeProperty = DependencyProperty.Register(
        nameof(EdgeSize), typeof(double), typeof(WorkshopSurface),
        new FrameworkPropertyMetadata(24.0, FrameworkPropertyMetadataOptions.AffectsRender),
        value => value is double size && double.IsFinite(size) && size >= 0);
    public double EdgeSize { get => (double)GetValue(EdgeSizeProperty); set => SetValue(EdgeSizeProperty, value); }

    protected override void OnRender(DrawingContext drawingContext)
    {
        base.OnRender(drawingContext);
        double width = ActualWidth, height = ActualHeight;
        if (width <= 0 || height <= 0) return;
        var slices = WorkshopArt.GetSurface(Material);
        double edgeX = Math.Min(EdgeSize, width / 2), edgeY = Math.Min(EdgeSize, height / 2);
        double[] xs = { 0, edgeX, width - edgeX, width };
        double[] ys = { 0, edgeY, height - edgeY, height };
        for (int row = 0; row < 3; row++)
            for (int column = 0; column < 3; column++)
                if (xs[column + 1] > xs[column] && ys[row + 1] > ys[row])
                    drawingContext.DrawImage(slices[row * 3 + column],
                        new Rect(xs[column], ys[row], xs[column + 1] - xs[column], ys[row + 1] - ys[row]));
    }
}

public sealed class WorkshopIcon : Image
{
    public static readonly DependencyProperty IconIndexProperty = DependencyProperty.Register(
        nameof(IconIndex), typeof(int), typeof(WorkshopIcon),
        new FrameworkPropertyMetadata(-1, (target, args) =>
            ((WorkshopIcon)target).Source = WorkshopArt.GetIconSource((int)args.NewValue)));
    public int IconIndex { get => (int)GetValue(IconIndexProperty); set => SetValue(IconIndexProperty, value); }
}
