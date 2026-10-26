using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace PaperSwitch.Controls;

/// <summary>功能圖示採原生向量輪廓，顏色繼承按鈕文字，裝飾插畫仍使用原圖集。</summary>
public sealed class ActionIcon : FrameworkElement
{
    public static readonly DependencyProperty IconIndexProperty = DependencyProperty.Register(
        nameof(IconIndex), typeof(int), typeof(ActionIcon),
        new FrameworkPropertyMetadata(-1, FrameworkPropertyMetadataOptions.AffectsRender));
    public int IconIndex { get => (int)GetValue(IconIndexProperty); set => SetValue(IconIndexProperty, value); }

    private static readonly Geometry[] Shapes = Array.ConvertAll(new[]
    {
        "M4,3 L14,3 19,8 19,21 4,21 Z M14,3 L14,8 19,8 M8,14 L15,14 M11.5,10.5 L11.5,17.5", // 加入
        "M4,3 L15,3 19,7 19,21 4,21 Z M8,12 L15,12 M8,16 L15,16",
        "M3,3 L12,3 12,15 3,15 Z M16,9 L21,9 21,21 11,21 11,18",
        "M3,4 L21,4 21,20 3,20 Z M4,17 L9,11 14,16 17,12 21,17 M15,8 L16,8",
        "M8,5 L3,9 8,13 M3,9 L14,9 C22,9 22,20 14,20", // 復原
        "M16,5 L21,9 16,13 M21,9 L10,9 C2,9 2,20 10,20", // 重做
        "M4,17 L16,5 20,9 8,21 4,21 Z M14,7 L18,11",
        "M5,3 L14,3 19,8 19,21 5,21 Z M14,3 L14,8 19,8",
        "M3,7 L3,4 10,4 13,7 21,7 21,20 3,20 Z",
        "M4,6 L20,6 M9,6 L9,3 15,3 15,6 M6,6 L7,21 17,21 18,6 M10,10 L10,17 M14,10 L14,17",
        "M20,8 C16,1 6,2 3,10 M3,5 L3,10 8,10 M4,16 C8,23 18,22 21,14 M16,14 L21,14 21,19",
        "M3,3 L21,3 21,21 3,21 Z M7,12 L11,16 18,8", // 全選
        "M20,12 L4,12 M10,6 L4,12 10,18",
        "M4,12 L20,12 M14,6 L20,12 14,18",
        "M3,3 L21,3 21,21 3,21 Z M12,3 L12,21 M12,3 L21,12 M12,12 L21,21", // 反選
        "M7,4 A9,9 0 1 0 20,10 M3,3 L3,9 9,9", // 逆轉
        "M17,4 A9,9 0 1 1 4,10 M21,3 L21,9 15,9" // 順轉
    }, value => { var geometry = Geometry.Parse(value); geometry.Freeze(); return geometry; });

    protected override void OnRender(DrawingContext context)
    {
        if (IconIndex < 0 || IconIndex >= Shapes.Length) return;
        var color = TextElement.GetForeground(this);
        var pen = new Pen(color, 1.8) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round, LineJoin = PenLineJoin.Round };
        var scale = Math.Min(ActualWidth, ActualHeight) / 24;
        context.PushTransform(new TranslateTransform((ActualWidth - 24 * scale) / 2, (ActualHeight - 24 * scale) / 2));
        context.PushTransform(new ScaleTransform(scale, scale));
        context.DrawGeometry(null, pen, Shapes[IconIndex]);
        context.Pop();
        context.Pop();
    }
}
