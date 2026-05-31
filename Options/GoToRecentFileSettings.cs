using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Settings;

namespace GoToRecentFile.Options
{
    /// <summary>
    /// Represents a single configurable setting entry displayed in the settings grid.
    /// </summary>
    internal sealed class SettingEntry : INotifyPropertyChanged
    {
        public string Name { get; }
        public string Type { get; }
        public string Default { get; }
        public string Description { get; }

        private string _value;
        public string Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    _value = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsModified)));
                }
            }
        }

        /// <summary>
        /// True when the current value differs from the default.
        /// </summary>
        public bool IsModified => !string.Equals(Value, Default, StringComparison.Ordinal);

        public event PropertyChangedEventHandler PropertyChanged;

        public SettingEntry(string name, string type, string defaultValue, string currentValue, string description)
        {
            Name = name;
            Type = type;
            Default = defaultValue;
            _value = currentValue;
            Description = description;
        }
    }

    /// <summary>
    /// Manages all configurable settings for the Go To Recent File extension.
    /// Settings are stored in the VS WritableSettingsStore and exposed as a flat list
    /// for display in a grid-based options page.
    /// </summary>
    internal static class GoToRecentFileSettings
    {
        private const string CollectionName = "GoToRecentFile\\Settings";

        // --- Setting Definitions (Name, Type, Default) ---
        // Window
        public const string WindowMinWidth = "WindowMinWidth";
        public const string WindowMinHeight = "WindowMinHeight";
        public const string WindowDefaultWidth = "WindowDefaultWidth";
        public const string WindowDefaultHeight = "WindowDefaultHeight";
        public const string WindowFontSize = "WindowFontSize";

        // Search
        public const string SearchDebounceMs = "SearchDebounceMs";
        public const string SearchPlaceholderText = "SearchPlaceholderText";

        // Highlighting
        public const string HighlightColor = "HighlightColor";

        // ListView
        public const string ColumnFileWidth = "ColumnFileWidth";
        public const string ColumnProjectWidth = "ColumnProjectWidth";
        public const string ColumnPathWidth = "ColumnPathWidth";
        public const string ColumnModifiedWidth = "ColumnModifiedWidth";

        // ListView Appearance
        public const string EnableAlternatingRowBackground = "EnableAlternatingRowBackground";
        public const string AlternateRowOpacity = "AlternateRowOpacity";

        // Light Theme Colors
        public const string LightSelectedRowBackground = "LightSelectedRowBackground";
        public const string LightSelectedRowBorder = "LightSelectedRowBorder";
        public const string LightSelectedRowForeground = "LightSelectedRowForeground";
        public const string LightAlternateRowBackground = "LightAlternateRowBackground";
        public const string LightColumnHeaderHoverBackground = "LightColumnHeaderHoverBackground";
        public const string LightColumnHeaderPressedBackground = "LightColumnHeaderPressedBackground";

        // Dark Theme Colors
        public const string DarkSelectedRowBackground = "DarkSelectedRowBackground";
        public const string DarkSelectedRowBorder = "DarkSelectedRowBorder";
        public const string DarkSelectedRowForeground = "DarkSelectedRowForeground";
        public const string DarkAlternateRowBackground = "DarkAlternateRowBackground";
        public const string DarkColumnHeaderHoverBackground = "DarkColumnHeaderHoverBackground";
        public const string DarkColumnHeaderPressedBackground = "DarkColumnHeaderPressedBackground";

        // Behavior
        public const string RememberSearchText = "RememberSearchText";
        public const string RememberWindowPosition = "RememberWindowPosition";
        public const string RememberWindowSize = "RememberWindowSize";
        public const string RememberColumnWidths = "RememberColumnWidths";

        /// <summary>
        /// Defines all settings with their type, default value, and description.
        /// </summary>
        private static readonly List<(string Name, string Type, string Default, string Description)> _definitions =
            new List<(string, string, string, string)>
            {
                // Window
                (WindowMinWidth,            "Integer", "760",
                    "Minimum width (in pixels) of the Go To Recent File dialog window. The window cannot be resized smaller than this."),
                (WindowMinHeight,           "Integer", "420",
                    "Minimum height (in pixels) of the Go To Recent File dialog window. The window cannot be resized smaller than this."),
                (WindowDefaultWidth,        "Integer", "980",
                    "Default width (in pixels) of the dialog window when opened for the first time or after resetting."),
                (WindowDefaultHeight,       "Integer", "560",
                    "Default height (in pixels) of the dialog window when opened for the first time or after resetting."),
                (WindowFontSize,            "Integer", "12",
                    "Font size (in points) used for text displayed in the dialog window."),

                // Search
                (SearchDebounceMs,          "Integer", "100",
                    "Delay in milliseconds before the search filter is applied after typing. Increase this value if search feels sluggish with large file lists."),
                (SearchPlaceholderText,     "String",  "Search filename (use P:project, F:path to search other columns)",
                    "Placeholder (watermark) text shown in the search box when it is empty. Helps users understand the search syntax."),

                // Highlighting
                (HighlightColor,            "String",  "Yellow",
                    "Background color used to highlight matching search terms in the file list. Use a named color (e.g., Yellow) or hex value (e.g., #FFFF00)."),

                // ListView columns
                (ColumnFileWidth,           "Integer", "200",
                    "Default width (in pixels) of the 'File' column showing the filename."),
                (ColumnProjectWidth,        "Integer", "130",
                    "Default width (in pixels) of the 'Project' column showing the project name."),
                (ColumnPathWidth,           "Integer", "440",
                    "Default width (in pixels) of the 'Path' column showing the full file path."),
                (ColumnModifiedWidth,       "Integer", "120",
                    "Default width (in pixels) of the 'Modified' column showing the last modified date."),

                // ListView Appearance
                (EnableAlternatingRowBackground, "Boolean", "True",
                    "When True, every other row in the file list uses a slightly different background color for easier readability."),
                (AlternateRowOpacity,       "String",  "08",
                    "Hex opacity value (00-FF) for the alternating row background. Lower values are more subtle. Example: 08 = ~3% opacity, 1A = ~10% opacity."),

                // Light Theme Colors
                (LightSelectedRowBackground,       "String", "#FFF0F0F0,#FFE0E0E0,#FFD0D0D0",
                    "Gradient colors for the selected row background in Light theme. Provide 3 comma-separated #AARRGGBB hex colors for top, middle, and bottom gradient stops."),
                (LightSelectedRowBorder,           "String", "#FF808080",
                    "Border color (#AARRGGBB hex) of the selected row in Light theme."),
                (LightSelectedRowForeground,       "String", "#FF000000",
                    "Text color (#AARRGGBB hex) of the selected row in Light theme."),
                (LightAlternateRowBackground,      "String", "#08000000",
                    "Background color (#AARRGGBB hex) of alternating rows in Light theme. Use a low alpha for subtlety (e.g., #08000000 = black at ~3% opacity)."),
                (LightColumnHeaderHoverBackground, "String", "#FFE8E8E8",
                    "Background color (#AARRGGBB hex) of column headers when hovered in Light theme."),
                (LightColumnHeaderPressedBackground, "String", "#FFD0D0D0",
                    "Background color (#AARRGGBB hex) of column headers when pressed/clicked in Light theme."),

                // Dark Theme Colors
                (DarkSelectedRowBackground,        "String", "#FF585858,#FF484848,#FF3A3A3A",
                    "Gradient colors for the selected row background in Dark theme. Provide 3 comma-separated #AARRGGBB hex colors for top, middle, and bottom gradient stops."),
                (DarkSelectedRowBorder,            "String", "#FF808080",
                    "Border color (#AARRGGBB hex) of the selected row in Dark theme."),
                (DarkSelectedRowForeground,        "String", "#FFFFFFFF",
                    "Text color (#AARRGGBB hex) of the selected row in Dark theme."),
                (DarkAlternateRowBackground,       "String", "#08FFFFFF",
                    "Background color (#AARRGGBB hex) of alternating rows in Dark theme. Use a low alpha for subtlety (e.g., #08FFFFFF = white at ~3% opacity)."),
                (DarkColumnHeaderHoverBackground,  "String", "#FF505050",
                    "Background color (#AARRGGBB hex) of column headers when hovered in Dark theme."),
                (DarkColumnHeaderPressedBackground,"String", "#FF404040",
                    "Background color (#AARRGGBB hex) of column headers when pressed/clicked in Dark theme."),

                // Behavior
                (RememberSearchText,        "Boolean", "True",
                    "When True, the search text is preserved between dialog invocations so you see the same filter when reopening."),
                (RememberWindowPosition,    "Boolean", "True",
                    "When True, the dialog remembers its last screen position and reopens there instead of centering on the owner window."),
                (RememberWindowSize,        "Boolean", "True",
                    "When True, the dialog remembers its last width and height and reopens at that size."),
                (RememberColumnWidths,      "Boolean", "True",
                    "When True, the column widths in the file list are preserved between dialog invocations."),
            };

        /// <summary>
        /// Loads all settings from the VS settings store and returns them as a list
        /// of <see cref="SettingEntry"/> for display in the grid.
        /// </summary>
        public static List<SettingEntry> Load()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var entries = new List<SettingEntry>();

            try
            {
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var store = settingsManager.GetReadOnlySettingsStore(SettingsScope.UserSettings);
                bool exists = store.CollectionExists(CollectionName);

                foreach (var (name, type, defaultValue, description) in _definitions)
                {
                    string current = defaultValue;
                    if (exists && store.PropertyExists(CollectionName, name))
                    {
                        current = store.GetString(CollectionName, name);
                    }
                    entries.Add(new SettingEntry(name, type, defaultValue, current, description));
                }
            }
            catch
            {
                // On failure, return defaults
                foreach (var (name, type, defaultValue, description) in _definitions)
                {
                    entries.Add(new SettingEntry(name, type, defaultValue, defaultValue, description));
                }
            }

            return entries;
        }

        /// <summary>
        /// Saves the provided setting entries to the VS settings store.
        /// </summary>
        public static void Save(IEnumerable<SettingEntry> entries)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var store = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

                if (!store.CollectionExists(CollectionName))
                    store.CreateCollection(CollectionName);

                foreach (var entry in entries)
                {
                    store.SetString(CollectionName, entry.Name, entry.Value);
                }
            }
            catch
            {
                // Silently ignore write failures.
            }
        }

        /// <summary>
        /// Gets a single setting value (current) from the VS settings store, or the default if not set.
        /// </summary>
        public static string GetValue(string settingName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string defaultValue = ""; 
            foreach (var (name, type, def, _) in _definitions)
            {
                if (string.Equals(name, settingName, StringComparison.Ordinal))
                {
                    defaultValue = def;
                    break;
                }
            }

            try
            {
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var store = settingsManager.GetReadOnlySettingsStore(SettingsScope.UserSettings);

                if (store.CollectionExists(CollectionName) && store.PropertyExists(CollectionName, settingName))
                    return store.GetString(CollectionName, settingName);
            }
            catch
            {
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets a setting as an integer, falling back to the default.
        /// </summary>
        public static int GetInt(string settingName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string val = GetValue(settingName);
            return int.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)
                ? result : 0;
        }

        /// <summary>
        /// Gets a setting as a boolean, falling back to the default.
        /// </summary>
        public static bool GetBool(string settingName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string val = GetValue(settingName);
            return string.Equals(val, "True", StringComparison.OrdinalIgnoreCase);
        }
    }
}
