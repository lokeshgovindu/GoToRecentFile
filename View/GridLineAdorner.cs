using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.VisualStudio.Shell;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Adorner that draws vertical grid lines over a ListView, aligned to column boundaries,
    /// from the header bottom to the control bottom.
    /// </summary>
    internal sealed class GridLineAdorner : Adorner
    {
        private readonly ListView _listView;

        public GridLineAdorner(ListView listView) : base(listView)
        {
            _listView = listView;
            IsHitTestVisible = false;
        }

        private Pen CreatePen()
        {
            // Derive a subtle separator color: blend the window background and window text
            // at 20% opacity so the line is always visible but never harsh in any VS theme.
            var bgBrush  = _listView.TryFindResource(VsBrushes.WindowKey)     as SolidColorBrush;
            var fgBrush  = _listView.TryFindResource(VsBrushes.WindowTextKey) as SolidColorBrush;

            Color bg = bgBrush?.Color ?? Colors.White;
            Color fg = fgBrush?.Color ?? Colors.Black;

            // Mix: 80% background + 20% foreground
            Color lineColor = Color.FromArgb(
                255,
                (byte)(bg.R * 0.80 + fg.R * 0.20),
                (byte)(bg.G * 0.80 + fg.G * 0.20),
                (byte)(bg.B * 0.80 + fg.B * 0.20));

            var brush = new SolidColorBrush(lineColor);
            brush.Freeze();
            var pen = new Pen(brush, 1);
            pen.Freeze();
            return pen;
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            if (!(_listView.View is GridView gridView))
                return;

            // Find the header row presenter to get actual column positions
            var headerPresenter = FindVisualChild<GridViewHeaderRowPresenter>(_listView);
            if (headerPresenter == null)
                return;

            double headerBottom = headerPresenter.TransformToAncestor(_listView)
                .Transform(new Point(0, headerPresenter.ActualHeight)).Y;
            double totalHeight = _listView.ActualHeight;

            Pen linePen = CreatePen();

            // Read the actual right edge of each column header for pixel-perfect alignment
            double x = headerPresenter.TransformToAncestor(_listView).Transform(new Point(0, 0)).X;
            foreach (var column in gridView.Columns)
            {
                x += column.ActualWidth;
                double snappedX = SnapToPixel(x) - 1;

                drawingContext.DrawLine(linePen,
                    new Point(snappedX, headerBottom),
                    new Point(snappedX, totalHeight - 1));
            }
        }

        private static double SnapToPixel(double value)
        {
            return System.Math.Round(value) + 0.5;
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                    return result;

                var descendant = FindVisualChild<T>(child);
                if (descendant != null)
                    return descendant;
            }

            return null;
        }
    }
}
