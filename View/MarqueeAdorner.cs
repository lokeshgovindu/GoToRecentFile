using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Adorner that renders the rubber-band selection rectangle over a ListView during a marquee drag.
    /// </summary>
    internal sealed class MarqueeAdorner : Adorner
    {
        private static readonly Brush FillBrush;
        private static readonly Brush BorderBrush;

        private Rect _rect;

        static MarqueeAdorner()
        {
            FillBrush = new SolidColorBrush(Color.FromArgb(40, 51, 153, 255));
            FillBrush.Freeze();

            BorderBrush = new SolidColorBrush(Color.FromArgb(180, 51, 153, 255));
            BorderBrush.Freeze();
        }

        public MarqueeAdorner(UIElement adornedElement) : base(adornedElement)
        {
            IsHitTestVisible = false;
            // Disable anti-aliasing so the border renders as exactly 1 device pixel.
            RenderOptions.SetEdgeMode(this, EdgeMode.Aliased);
        }

        public void Update(Point anchor, Point current)
        {
            _rect = new Rect(anchor, current);
            InvalidateVisual();
        }

        public void Clear()
        {
            _rect = Rect.Empty;
            InvalidateVisual();
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            if (_rect.IsEmpty || (_rect.Width < 1 && _rect.Height < 1))
                return;

            // Resolve 1 device pixel in logical units at the current DPI.
            Matrix m = PresentationSource.FromVisual(this)?.CompositionTarget?.TransformToDevice ?? Matrix.Identity;
            double px = m.M11 > 0 ? 1.0 / m.M11 : 1.0;
            double py = m.M22 > 0 ? 1.0 / m.M22 : 1.0;

            // Snap each edge to the nearest device pixel boundary.
            double x0 = System.Math.Round(_rect.Left  / px) * px;
            double y0 = System.Math.Round(_rect.Top   / py) * py;
            double x1 = System.Math.Round(_rect.Right / px) * px;
            double y1 = System.Math.Round(_rect.Bottom/ py) * py;
            Rect snapped = new Rect(new Point(x0, y0), new Point(x1, y1));

            // Pen thickness is exactly 1 device pixel.
            var pen = new Pen(BorderBrush, px);
            pen.Freeze();

            drawingContext.DrawRectangle(FillBrush, pen, snapped);
        }
    }
}
