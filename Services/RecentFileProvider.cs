using System.Collections.Generic;
using System.IO;
using EnvDTE;
using EnvDTE80;
using GoToRecentFile.Models;

namespace GoToRecentFile
{
    /// <summary>
    /// Retrieves the list of recently opened files ordered by most-recently-activated (Ctrl+Tab order).
    /// </summary>
    internal static class RecentFileProvider
    {
        /// <summary>
        /// Gets the recently opened documents as <see cref="RecentFileEntry"/> in MRU order.
        /// </summary>
        public static async System.Threading.Tasks.Task<IReadOnlyList<RecentFileEntry>> GetRecentFilesAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = await VS.GetServiceAsync<DTE, DTE2>();
            var entries = new List<RecentFileEntry>();

            foreach (EnvDTE.Window window in dte.Windows)
            {
                if (window.Kind != "Document")
                    continue;

                try
                {
                    Document doc = window.Document;
                    if (doc == null)
                        continue;

                    string path = doc.FullName;
                    if (string.IsNullOrEmpty(path) || !File.Exists(path))
                        continue;

                    string projectName = "";
                    try
                    {
                        projectName = doc.ProjectItem?.ContainingProject?.Name ?? "";
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                    }

                    var modified = File.GetLastWriteTime(path);
                    entries.Add(new RecentFileEntry(path, projectName, modified));
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Window may not have an associated document
                }
            }

            return entries;
        }
    }
}
