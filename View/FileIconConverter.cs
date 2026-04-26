using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Converts a file path to its associated shell icon ImageSource.
    /// </summary>
    internal sealed class FileIconConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string path = value as string;
            if (string.IsNullOrEmpty(path))
                return null;

            string ext = System.IO.Path.GetExtension(path);
            return FileIconHelper.GetIconForExtension(ext);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }
}
