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
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (EnvDTE.Window window in dte.Windows)
            {
                if (window.Kind != "Document")
                    continue;

                try
                {
                    string path = null;
                    string projectName = "";

                    Document doc = window.Document;
                    if (doc != null)
                    {
                        path = doc.FullName;
                        try
                        {
                            projectName = doc.ProjectItem?.ContainingProject?.Name ?? "";
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                        }
                    }
                    else
                    {
                        // Document property can be null for open but not-yet-loaded tabs
                        try
                        {
                            ProjectItem pi = window.ProjectItem;
                            if (pi != null)
                            {
                                path = pi.FileNames[1];
                                try
                                {
                                    projectName = pi.ContainingProject?.Name ?? "";
                                }
                                catch (System.Runtime.InteropServices.COMException)
                                {
                                }
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                        }
                    }

                    if (string.IsNullOrEmpty(path) || !File.Exists(path))
                        continue;

                    if (!seen.Add(path))
                        continue;

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
