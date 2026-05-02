using System.Windows;
using System.Windows.Interop;
using EnvDTE;
using EnvDTE80;
using GoToRecentFile.View;

namespace GoToRecentFile
{
    [Command(PackageIds.GoToRecentFile)]
    internal sealed class GoToRecentFileCommand : BaseCommand<GoToRecentFileCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var recentFiles = await RecentFileProvider.GetRecentFilesAsync();

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = await VS.GetServiceAsync<DTE, DTE2>();

            var dialog = new GoToRecentFileWindow(recentFiles);

            // Set the VS main window as owner so it behaves as a proper modal dialog
            var hwnd = (IntPtr)dte.MainWindow.HWnd;
            var helper = new WindowInteropHelper(dialog) { Owner = hwnd };

            dialog.FileRemoved += path =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                try
                {
                    foreach (EnvDTE.Window window in dte.Windows)
                    {
                        if (window.Kind != "Document")
                            continue;

                        Document doc = window.Document;
                        if (doc != null && string.Equals(doc.FullName, path, System.StringComparison.OrdinalIgnoreCase))
                        {
                            window.Close(vsSaveChanges.vsSaveChangesPrompt);
                            break;
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                }
            };

            bool? result = dialog.ShowDialog();

            if (result == true && !string.IsNullOrEmpty(dialog.SelectedFilePath))
            {
                await VS.Documents.OpenAsync(dialog.SelectedFilePath);
            }
        }
    }
}
