using System;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Xunit;

namespace PaperSwitch.Tests
{
    public class ScrollBarHitTests
    {
        [Fact]
        public void IsInsideScrollBar_ShouldRecognizeScrollBarParts()
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var bar = new ScrollBar();
                    Assert.True(MainWindow.IsInsideScrollBar(bar));
                    Assert.False(MainWindow.IsInsideScrollBar(new Border()));
                    Assert.False(MainWindow.IsInsideScrollBar(null));
                }
                catch (Exception ex) { failure = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) throw failure;
        }
    }
}
