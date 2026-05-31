using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Settings;
using System.Linq;

namespace GoToRecentFile.Services
{
    /// <summary>
    /// Persists window layout settings (position, size, column widths) across VS sessions
    /// using the VS WritableSettingsStore.
    /// </summary>
    internal static class WindowLayoutSettings
    {
        private const string CollectionName = "GoToRecentFile";
        private const string LeftKey = "WindowLeft";
        private const string TopKey = "WindowTop";
        private const string WidthKey = "WindowWidth";
        private const string HeightKey = "WindowHeight";
        private const string ColumnWidthsKey = "ColumnWidths";

        public static double? WindowLeft { get; set; }
        public static double? WindowTop { get; set; }
        public static double? WindowWidth { get; set; }
        public static double? WindowHeight { get; set; }
        public static double[] ColumnWidths { get; set; }

        /// <summary>
        /// Loads persisted layout values from the VS settings store.
        /// Must be called on the UI thread.
        /// </summary>
        public static void Load()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var store = settingsManager.GetReadOnlySettingsStore(SettingsScope.UserSettings);

                if (!store.CollectionExists(CollectionName))
                    return;

                if (store.PropertyExists(CollectionName, LeftKey))
                    WindowLeft = store.GetString(CollectionName, LeftKey) is string s && double.TryParse(s, out double v) ? v : (double?)null;

                if (store.PropertyExists(CollectionName, TopKey))
                    WindowTop = store.GetString(CollectionName, TopKey) is string s && double.TryParse(s, out double v) ? v : (double?)null;

                if (store.PropertyExists(CollectionName, WidthKey))
                    WindowWidth = store.GetString(CollectionName, WidthKey) is string s && double.TryParse(s, out double v) ? v : (double?)null;

                if (store.PropertyExists(CollectionName, HeightKey))
                    WindowHeight = store.GetString(CollectionName, HeightKey) is string s && double.TryParse(s, out double v) ? v : (double?)null;

                if (store.PropertyExists(CollectionName, ColumnWidthsKey))
                {
                    string raw = store.GetString(CollectionName, ColumnWidthsKey);
                    if (!string.IsNullOrEmpty(raw))
                    {
                        ColumnWidths = raw.Split(',')
                            .Select(p => double.TryParse(p, out double d) ? d : 0)
                            .ToArray();
                    }
                }
            }
            catch
            {
                // Silently ignore settings read failures.
            }
        }

        /// <summary>
        /// Saves current layout values to the VS settings store.
        /// Must be called on the UI thread.
        /// </summary>
        public static void Save()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var store = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

                if (!store.CollectionExists(CollectionName))
                    store.CreateCollection(CollectionName);

                if (WindowLeft.HasValue)
                    store.SetString(CollectionName, LeftKey, WindowLeft.Value.ToString());

                if (WindowTop.HasValue)
                    store.SetString(CollectionName, TopKey, WindowTop.Value.ToString());

                if (WindowWidth.HasValue)
                    store.SetString(CollectionName, WidthKey, WindowWidth.Value.ToString());

                if (WindowHeight.HasValue)
                    store.SetString(CollectionName, HeightKey, WindowHeight.Value.ToString());

                if (ColumnWidths != null && ColumnWidths.Length > 0)
                    store.SetString(CollectionName, ColumnWidthsKey, string.Join(",", ColumnWidths));
            }
            catch
            {
                // Silently ignore settings write failures.
            }
        }
    }
}
