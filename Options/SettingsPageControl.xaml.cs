using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;

namespace GoToRecentFile.Options
{
    public partial class SettingsPageControl : UserControl
    {
        private List<SettingEntry> _allSettings;

        public SettingsPageControl()
        {
            InitializeComponent();
            ThreadHelper.ThrowIfNotOnUIThread();
            LoadSettings();
        }

        internal void LoadSettings()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _allSettings = GoToRecentFileSettings.Load();
            SettingsGrid.ItemsSource = _allSettings;
        }

        internal void SaveSettings()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_allSettings != null)
            {
                GoToRecentFileSettings.Save(_allSettings);
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = SearchBox.Text?.Trim();
            if (string.IsNullOrEmpty(filter))
            {
                SettingsGrid.ItemsSource = _allSettings;
            }
            else
            {
                SettingsGrid.ItemsSource = _allSettings
                    .Where(s => s.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                s.Value.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            SearchPlaceholder.Visibility = string.IsNullOrEmpty(SearchBox.Text)
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        private void ResetAll_Click(object sender, RoutedEventArgs e)
        {
            if (_allSettings == null) return;

            var result = System.Windows.MessageBox.Show(
                "Reset all settings to their default values?",
                "Reset Settings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                foreach (var entry in _allSettings)
                {
                    entry.Value = entry.Default;
                }
                SettingsGrid.Items.Refresh();
                UpdateApplyButtonState();
            }
        }

        private void ResetSelected_Click(object sender, RoutedEventArgs e)
        {
            if (SettingsGrid.SelectedItem is SettingEntry entry)
            {
                entry.Value = entry.Default;
                SettingsGrid.Items.Refresh();
                UpdateApplyButtonState();
            }
        }

        private void SettingsGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (SettingsGrid.SelectedItem is SettingEntry entry)
            {
                HelpTextBox.Text = entry.Description;
            }
            else
            {
                HelpTextBox.Text = string.Empty;
            }
        }

        private void SettingsGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (!(SettingsGrid.SelectedItem is SettingEntry entry)) return;

            var dialog = new SettingEditorDialog(entry.Name, entry.Value, entry.Type);
            dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            // Parent the dialog to the VS main window to avoid modality/focus issues.
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var shell = (Microsoft.VisualStudio.Shell.Interop.IVsUIShell)ServiceProvider.GlobalProvider
                    .GetService(typeof(Microsoft.VisualStudio.Shell.Interop.SVsUIShell));
                if (shell != null)
                {
                    shell.GetDialogOwnerHwnd(out IntPtr hwnd);
                    if (hwnd != IntPtr.Zero)
                    {
                        new WindowInteropHelper(dialog).Owner = hwnd;
                        dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                    }
                }
            }
            catch { }

            if (dialog.ShowDialog() == true)
            {
                entry.Value = dialog.SettingValue;
                SettingsGrid.Items.Refresh();
                UpdateApplyButtonState();
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SaveSettings();
            ApplyButton.IsEnabled = false;
        }

        private void UpdateApplyButtonState()
        {
            ApplyButton.IsEnabled = _allSettings != null && _allSettings.Any(s => s.IsModified);
        }
    }

    /// <summary>
    /// Converts text length to visibility: Visible when length is 0 (shows placeholder), Collapsed otherwise.
    /// </summary>
    public class InverseBoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int length)
                return length == 0 ? Visibility.Visible : Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
