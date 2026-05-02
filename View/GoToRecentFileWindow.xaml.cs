using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using GoToRecentFile.Models;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Dialog that displays recently opened files with search filtering.
    /// </summary>
    public partial class GoToRecentFileWindow : Window
    {
        private readonly List<RecentFileEntry> _allFiles;
        private GridViewColumnHeader _lastHeaderClicked;
        private ListSortDirection _lastDirection = ListSortDirection.Ascending;

        // Persisted position across dialog invocations within the same VS session.
        private static double? _lastLeft;
        private static double? _lastTop;

        /// <summary>
        /// Dependency property exposing the current search words for the highlight converter binding.
        /// </summary>
        public static readonly DependencyProperty SearchWordsProperty =
            DependencyProperty.Register(nameof(SearchWords), typeof(string[]), typeof(GoToRecentFileWindow),
                new PropertyMetadata(Array.Empty<string>()));

        /// <summary>
        /// Gets or sets the current search words used for highlighting.
        /// </summary>
        public string[] SearchWords
        {
            get { return (string[])GetValue(SearchWordsProperty); }
            set { SetValue(SearchWordsProperty, value); }
        }

        /// <summary>
        /// Gets the file path the user selected, or null if cancelled.
        /// </summary>
        public string SelectedFilePath { get; private set; }

        /// <summary>
        /// Raised when the user removes a file from the list via the close button.
        /// The argument is the full path of the removed file.
        /// </summary>
        internal event Action<string> FileRemoved;

        private GridLineAdorner _gridLineAdorner;

        #region Win32 – title-bar button customisation

        private const int GWL_STYLE   = -16;
        private const int GWL_EXSTYLE = -20;
        private const int WS_MINIMIZEBOX  = 0x00020000;
        private const int WS_MAXIMIZEBOX  = 0x00010000;
        private const int WS_EX_CONTEXTHELP = 0x00000400;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            IntPtr hwnd = new WindowInteropHelper(this).Handle;

            // Remove minimize and maximize buttons
            int style = GetWindowLong(hwnd, GWL_STYLE);
            style &= ~WS_MINIMIZEBOX;
            style &= ~WS_MAXIMIZEBOX;
            SetWindowLong(hwnd, GWL_STYLE, style);

            // Add the Help (?) button
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            exStyle |= WS_EX_CONTEXTHELP;
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);
        }

        #endregion

        internal GoToRecentFileWindow(IReadOnlyList<RecentFileEntry> files)
        {
            _allFiles = files?.ToList() ?? new List<RecentFileEntry>();

            InitializeComponent();

            SourceInitialized += OnSourceInitialized;

            FileListView.ItemsSource = _allFiles;
            UpdateStatus(_allFiles.Count, _allFiles.Count);

            if (_allFiles.Count > 0)
            {
                FileListView.SelectedIndex = 0;
            }

            Loaded += (s, e) =>
            {
                // Restore persisted position if available, otherwise CenterOwner is used.
                if (_lastLeft.HasValue && _lastTop.HasValue)
                {
                    WindowStartupLocation = WindowStartupLocation.Manual;
                    Left = _lastLeft.Value;
                    Top = _lastTop.Value;
                }

                SearchBox.Focus();
                AttachGridLineAdorner();
            };

            FileListView.SizeChanged += (s, e) => InvalidateGridLines();

            Closing += (s, e) =>
            {
                _lastLeft = Left;
                _lastTop = Top;
            };
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = SearchBox.Text?.Trim();

            if (string.IsNullOrEmpty(filter))
            {
                SearchWords = Array.Empty<string>();
                FileListView.ItemsSource = _allFiles;
                UpdateStatus(_allFiles.Count, _allFiles.Count);
            }
            else
            {
                string[] words = filter.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                SearchWords = words;

                var filtered = _allFiles
                    .Where(f => words.All(w => f.FileName.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0))
                    .ToList();

                FileListView.ItemsSource = filtered;
                UpdateStatus(filtered.Count, _allFiles.Count);
            }

            if (FileListView.Items.Count > 0)
            {
                FileListView.SelectedIndex = 0;
            }
        }

        private void SearchBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                MoveSelection(1);
                e.Handled = true;
            }
            else if (e.Key == Key.Up)
            {
                MoveSelection(-1);
                e.Handled = true;
            }
            else if (e.Key == Key.PageDown)
            {
                MoveSelection(10);
                e.Handled = true;
            }
            else if (e.Key == Key.PageUp)
            {
                MoveSelection(-10);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                AcceptSelection();
                e.Handled = true;
            }
        }

        private void MoveSelection(int delta)
        {
            int count = FileListView.Items.Count;
            if (count == 0)
                return;

            int index = FileListView.SelectedIndex + delta;
            index = Math.Max(0, Math.Min(index, count - 1));
            FileListView.SelectedIndex = index;
            FileListView.ScrollIntoView(FileListView.SelectedItem);
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            AcceptSelection();
        }

        private void FileListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            AcceptSelection();
        }

        // Maps column header text to the sort property name.
        private static readonly Dictionary<string, string> _columnSortMap = new Dictionary<string, string>
        {
            { "File", "FileName" },
            { "Project", "Project" },
            { "Path", "FullPath" },
            { "Modified", "Modified" }
        };

        private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            if (!(e.OriginalSource is GridViewColumnHeader header) || header.Role == GridViewColumnHeaderRole.Padding)
                return;

            string headerText = header.Column.Header as string;
            if (headerText == null || !_columnSortMap.TryGetValue(headerText, out string sortBy))
                return;

            ListSortDirection direction = (header == _lastHeaderClicked && _lastDirection == ListSortDirection.Ascending)
                ? ListSortDirection.Descending
                : ListSortDirection.Ascending;

            ICollectionView view = CollectionViewSource.GetDefaultView(FileListView.ItemsSource);
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(sortBy, direction));
            view.Refresh();

            // Remove arrow from previous header
            if (_lastHeaderClicked != null && _lastHeaderClicked != header)
            {
                _lastHeaderClicked.Column.HeaderTemplate = null;
            }

            // Set arrow indicator on current header
            header.Column.HeaderTemplate = (DataTemplate)FindResource(
                direction == ListSortDirection.Ascending ? "HeaderTemplateArrowUp" : "HeaderTemplateArrowDown");

            _lastHeaderClicked = header;
            _lastDirection = direction;
        }

        private void CloseFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (!((sender as System.Windows.Controls.Button)?.Tag is RecentFileEntry entry))
                return;

            int selectedIndex = FileListView.SelectedIndex;
            _allFiles.Remove(entry);

            // Re-raise the event so the caller can close the document in VS
            FileRemoved?.Invoke(entry.FullPath);

            // Re-apply the current filter
            SearchBox_TextChanged(SearchBox, null);

            // Restore selection near the same position
            if (FileListView.Items.Count > 0)
            {
                FileListView.SelectedIndex = Math.Min(selectedIndex, FileListView.Items.Count - 1);
            }
        }

        private void AcceptSelection()
        {
            if (FileListView.SelectedItem is RecentFileEntry entry)
            {
                SelectedFilePath = entry.FullPath;
                DialogResult = true;
                Close();
            }
        }

        private void UpdateStatus(int shown, int total)
        {
            Title = $"Go To Recent File [{shown} of {total}]";
        }

        /// <summary>
        /// Attaches the grid line adorner to the ListView and listens for column width changes.
        /// </summary>
        private void AttachGridLineAdorner()
        {
            var adornerLayer = AdornerLayer.GetAdornerLayer(FileListView);
            if (adornerLayer == null)
                return;

            _gridLineAdorner = new GridLineAdorner(FileListView);
            adornerLayer.Add(_gridLineAdorner);

            // Listen for column width changes to redraw lines
            if (FileListView.View is GridView gridView)
            {
                foreach (var column in gridView.Columns)
                {
                    ((System.ComponentModel.INotifyPropertyChanged)column).PropertyChanged += (s, e) =>
                    {
                        if (e.PropertyName == nameof(GridViewColumn.ActualWidth))
                        {
                            InvalidateGridLines();
                        }
                    };
                }
            }
        }

        private void InvalidateGridLines()
        {
            _gridLineAdorner?.InvalidateVisual();
        }


    }
}
