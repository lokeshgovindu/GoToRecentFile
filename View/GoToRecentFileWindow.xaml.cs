using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
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
        // Prevents SearchBox_TextChanged from auto-selecting item 0 while RemoveFiles restores selection.
        private bool _restoringSelection;

        // Marquee (rubber-band) selection
        private MarqueeAdorner _marqueeAdorner;
        private Point _marqueeDragStart;
        private bool _marqueeActive;

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
                AttachMarqueeAdorner();
            };

            FileListView.SizeChanged += (s, e) => InvalidateGridLines();

            Closing += (s, e) =>
            {
                _lastLeft = Left;
                _lastTop = Top;
            };
        }

        private void FileListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            OpenButton.IsEnabled = FileListView.SelectedItems.Count == 1;
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

            if (!_restoringSelection && FileListView.Items.Count > 0)
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

            RemoveFiles(new List<RecentFileEntry> { entry });
        }

        /// <summary>
        /// Removes the given entries from the list, fires <see cref="FileRemoved"/> for each,
        /// and refreshes the view while preserving the remaining selection.
        /// </summary>
        private void RemoveFiles(IReadOnlyList<RecentFileEntry> entries)
        {
            if (entries.Count == 0)
                return;

            // Remember which items were selected but are NOT being removed.
            var keepSelected = FileListView.SelectedItems
                .Cast<RecentFileEntry>()
                .Except(entries)
                .ToList();

            int selectedIndex = FileListView.SelectedIndex;

            foreach (RecentFileEntry entry in entries)
            {
                _allFiles.Remove(entry);
                FileRemoved?.Invoke(entry.FullPath);
            }

            if (keepSelected.Count > 0)
            {
                _restoringSelection = true;
                FileListView.ItemsSource = null;
                SearchBox_TextChanged(SearchBox, null);
                _restoringSelection = false;

                // Containers are generated after ItemsSource is set; defer until layout is done.
#pragma warning disable VSSDK007
                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {
                    await System.Threading.Tasks.Task.Delay(0); // yield to allow layout pass
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    FileListView.SelectedItems.Clear();
                    foreach (RecentFileEntry item in keepSelected)
                    {
                        if (FileListView.ItemContainerGenerator.ContainerFromItem(item) is ListViewItem container)
                            container.IsSelected = true;
                    }
                }).FileAndForget(nameof(RemoveFiles));
#pragma warning restore VSSDK007
            }
            else
            {
                FileListView.ItemsSource = null;
                SearchBox_TextChanged(SearchBox, null);

                if (FileListView.Items.Count > 0)
                    FileListView.SelectedIndex = Math.Min(selectedIndex, FileListView.Items.Count - 1);
            }
        }

        #region Context menu

        private void FileListView_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            bool hasSelection = FileListView.SelectedItems.Count > 0;
            bool single = FileListView.SelectedItems.Count == 1;

            CtxCloseFiles.Header = FileListView.SelectedItems.Count > 1 ? "Close Files" : "Close File";
            CtxCloseFiles.IsEnabled = hasSelection;
            CtxCopyFilename.IsEnabled = single;
            CtxOpenContainingFolder.IsEnabled = single;
            CtxCopyFilePath.IsEnabled = single;

            string project = single
                ? (FileListView.SelectedItem as RecentFileEntry)?.Project
                : null;
            bool hasProject = !string.IsNullOrEmpty(project);
            CtxSelectProjectFiles.Header = hasProject ? $"Select Files in '{project}'" : "Select Files in Project";
            CtxSelectProjectFiles.IsEnabled = hasProject;

            CtxSelectMiscFiles.IsEnabled = FileListView.Items
                .OfType<RecentFileEntry>()
                .Any(f => string.Equals(f.Project, "Miscellaneous Files", StringComparison.OrdinalIgnoreCase));
        }

        private void CtxCloseFiles_Click(object sender, RoutedEventArgs e)
        {
            var toRemove = FileListView.SelectedItems.Cast<RecentFileEntry>().ToList();
            RemoveFiles(toRemove);
        }

        private void CtxCopyFilename_Click(object sender, RoutedEventArgs e)
        {
            if (FileListView.SelectedItem is RecentFileEntry entry)
            {
#pragma warning disable VSSDK007
                ThreadHelper.JoinableTaskFactory.RunAsync(() => TrySetClipboardAsync(entry.FileName)).FileAndForget(nameof(CtxCopyFilename_Click));
#pragma warning restore VSSDK007
            }
        }

        private void CtxOpenContainingFolder_Click(object sender, RoutedEventArgs e)
        {
            if (!(FileListView.SelectedItem is RecentFileEntry entry))
                return;

            string folder = Path.GetDirectoryName(entry.FullPath);
            if (Directory.Exists(folder))
                Process.Start("explorer.exe", $"/select,\"{entry.FullPath}\"");
        }

        private void CtxCopyFilePath_Click(object sender, RoutedEventArgs e)
        {
            if (FileListView.SelectedItem is RecentFileEntry entry)
            {
#pragma warning disable VSSDK007
                ThreadHelper.JoinableTaskFactory.RunAsync(() => TrySetClipboardAsync(entry.FullPath)).FileAndForget(nameof(CtxCopyFilePath_Click));
#pragma warning restore VSSDK007
            }
        }

        private static async System.Threading.Tasks.Task TrySetClipboardAsync(string text, int retries = 5)
        {
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    Clipboard.SetDataObject(text, true);
                    return;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                    when (ex.HResult == unchecked((int)0x800401D0)) // CLIPBRD_E_CANT_OPEN
                {
                    await System.Threading.Tasks.Task.Delay(10);
                }
            }
        }

        private void CtxSelectProjectFiles_Click(object sender, RoutedEventArgs e)
        {
            if (!(FileListView.SelectedItem is RecentFileEntry selected) ||
                string.IsNullOrEmpty(selected.Project))
                return;

            FileListView.SelectedItems.Clear();
            foreach (object item in FileListView.Items)
            {
                if (item is RecentFileEntry entry &&
                    string.Equals(entry.Project, selected.Project, StringComparison.OrdinalIgnoreCase))
                {
                    if (FileListView.ItemContainerGenerator.ContainerFromItem(entry) is ListViewItem container)
                        container.IsSelected = true;
                }
            }
        }

        private void CtxSelectMiscFiles_Click(object sender, RoutedEventArgs e)
        {
            FileListView.SelectedItems.Clear();
            foreach (object item in FileListView.Items)
            {
                if (item is RecentFileEntry entry &&
                    string.Equals(entry.Project, "Miscellaneous Files", StringComparison.OrdinalIgnoreCase))
                {
                    if (FileListView.ItemContainerGenerator.ContainerFromItem(entry) is ListViewItem container)
                        container.IsSelected = true;
                }
            }
        }

        private void CtxSelectAll_Click(object sender, RoutedEventArgs e)
        {
            FileListView.SelectAll();
        }

        #endregion

        private void FileListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.A && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                FileListView.SelectAll();
                e.Handled = true;
            }
            else if (e.Key == Key.Delete)
            {
                var toRemove = FileListView.SelectedItems.Cast<RecentFileEntry>().ToList();
                RemoveFiles(toRemove);
                e.Handled = true;
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

        private void AttachMarqueeAdorner()
        {
            var adornerLayer = AdornerLayer.GetAdornerLayer(FileListView);
            if (adornerLayer == null)
                return;

            _marqueeAdorner = new MarqueeAdorner(FileListView);
            adornerLayer.Add(_marqueeAdorner);

            FileListView.PreviewMouseLeftButtonDown += FileListView_MarqueeMouseDown;
            FileListView.PreviewMouseMove           += FileListView_MarqueeMouseMove;
            FileListView.PreviewMouseLeftButtonUp   += FileListView_MarqueeMouseUp;
        }

        private void FileListView_MarqueeMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Only start marquee when clicking on the empty area (not on an item).
            if (e.OriginalSource is DependencyObject source &&
                GetVisualAncestor<ListViewItem>(source) != null)
                return;

            _marqueeDragStart = e.GetPosition(FileListView);
            _marqueeActive = true;
            FileListView.CaptureMouse();

            FileListView.SelectedItems.Clear();
            e.Handled = false;
        }

        private void FileListView_MarqueeMouseMove(object sender, MouseEventArgs e)
        {
            if (!_marqueeActive || e.LeftButton != MouseButtonState.Pressed)
                return;

            Point current = e.GetPosition(FileListView);
            _marqueeAdorner.Update(_marqueeDragStart, current);

            UpdateMarqueeSelection(current);
        }

        private void FileListView_MarqueeMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_marqueeActive)
                return;

            _marqueeActive = false;
            _marqueeAdorner.Clear();
            FileListView.ReleaseMouseCapture();
        }

        private void UpdateMarqueeSelection(Point current)
        {
            Rect dragRect = new Rect(_marqueeDragStart, current);

            foreach (object item in FileListView.Items)
            {
                if (!(FileListView.ItemContainerGenerator.ContainerFromItem(item) is ListViewItem container))
                    continue;

                // Get bounding rect of the item relative to the ListView.
                Point topLeft = container.TranslatePoint(new Point(0, 0), FileListView);
                Rect itemRect = new Rect(topLeft, new Size(container.ActualWidth, container.ActualHeight));

                container.IsSelected = dragRect.IntersectsWith(itemRect);
            }
        }

        private static T GetVisualAncestor<T>(DependencyObject element) where T : DependencyObject
        {
            DependencyObject current = VisualTreeHelper.GetParent(element);
            while (current != null)
            {
                if (current is T match)
                    return match;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }


    }
}
