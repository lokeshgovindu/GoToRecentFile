using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Returns <see cref="Visibility.Visible"/> when the bound text length is zero,
    /// <see cref="Visibility.Collapsed"/> otherwise. Used for placeholder watermark text.
    /// </summary>
    internal sealed class PlaceholderVisibilityConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int length = value is int i ? i : 0;
            return length == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }
}
