using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Adorner that draws vertical grid lines over a ListView, aligned to column boundaries,
    /// from the header bottom to the control bottom.
    /// </summary>
    internal sealed class GridLineAdorner : Adorner
    {
        private readonly ListView _listView;
        private static readonly Pen LinePen = CreatePen();

        public GridLineAdorner(ListView listView) : base(listView)
        {
            _listView = listView;
            IsHitTestVisible = false;
        }

        private static Pen CreatePen()
        {
            var pen = new Pen(Brushes.LightGray, 1);
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

            // Read the actual right edge of each column header for pixel-perfect alignment
            double x = headerPresenter.TransformToAncestor(_listView).Transform(new Point(0, 0)).X;
            foreach (var column in gridView.Columns)
            {
                x += column.ActualWidth;
                double snappedX = SnapToPixel(x) - 1;

                drawingContext.DrawLine(LinePen,
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
