using System;
using System.Collections.Generic;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Retrieves Visual Studio image monikers for file extensions,
    /// producing the same icons shown in Solution Explorer.
    /// </summary>
    internal static class FileIconHelper
    {
        private static readonly Dictionary<string, ImageSource> _cache =
            new Dictionary<string, ImageSource>(StringComparer.OrdinalIgnoreCase);

        private static IVsImageService2 _imageService;

        /// <summary>
        /// Gets the VS ImageMoniker for the given file extension.
        /// </summary>
        public static ImageMoniker GetMonikerForExtension(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                extension = ".txt";

            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var svc = GetImageService();
                if (svc != null)
                {
                    return svc.GetImageMonikerForFile("file" + extension);
                }
            }
            catch (Exception)
            {
            }

            return default;
        }

        /// <summary>
        /// Gets the VS image for the given file extension as a WPF ImageSource, cached per extension.
        /// </summary>
        public static ImageSource GetIconForExtension(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                extension = ".txt";

            if (_cache.TryGetValue(extension, out ImageSource cached))
                return cached;

            ImageSource imageSource = null;

            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var imageService = GetImageService();
                if (imageService != null)
                {
                    ImageMoniker moniker = imageService.GetImageMonikerForFile("file" + extension);

                    if (moniker.Guid != Guid.Empty)
                    {
                        var attributes = new ImageAttributes
                        {
                            StructSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(ImageAttributes)),
                            ImageType = (uint)_UIImageType.IT_Bitmap,
                            Format = (uint)_UIDataFormat.DF_WPF,
                            LogicalWidth = 16,
                            LogicalHeight = 16,
                            Flags = unchecked((uint)_ImageAttributesFlags.IAF_RequiredFlags),
                            Background = 0x00FFFFFF // Transparent background
                        };

                        IVsUIObject uiObject = imageService.GetImage(moniker, attributes);
                        if (uiObject != null)
                        {
                            uiObject.get_Data(out object data);
                            if (data is BitmapSource bmp)
                            {
                                imageSource = bmp;
                                if (imageSource.CanFreeze)
                                    imageSource.Freeze();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Fall back gracefully if VS image service is unavailable
            }

            _cache[extension] = imageSource;
            return imageSource;
        }

        private static IVsImageService2 GetImageService()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_imageService == null)
            {
                _imageService = (IVsImageService2)Package.GetGlobalService(typeof(SVsImageService));
            }

            return _imageService;
        }
    }
}
