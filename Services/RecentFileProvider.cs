using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using EnvDTE;
using EnvDTE80;
using GoToRecentFile.Models;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

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

            // Build a map of project name -> color from the solution hierarchy order
            var projectColorMap = await GetProjectColorMapAsync();

            var entries = new List<RecentFileEntry>();
            var seen = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

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

                    SolidColorBrush brush = null;
                    if (!string.IsNullOrEmpty(projectName))
                    {
                        projectColorMap.TryGetValue(projectName, out brush);
                    }

                    entries.Add(new RecentFileEntry(path, projectName, modified, brush));
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Window may not have an associated document
                }
            }

            return entries;
        }

        /// <summary>
        /// VS tab "Project" colorization palette. These are the 10 colors VS cycles through
        /// for project-based tab colorization, assigned in solution project order.
        /// </summary>
        private static readonly Color[] VsProjectColorPalette = new[]
        {
            Color.FromRgb(0xCE, 0x91, 0xD8), // purple/magenta
            Color.FromRgb(0x56, 0xC2, 0x6D), // green
            Color.FromRgb(0xD8, 0x8C, 0x4E), // orange/brown
            Color.FromRgb(0x5D, 0xAE, 0xD6), // blue
            Color.FromRgb(0xD8, 0x56, 0x6C), // red/pink
            Color.FromRgb(0x8E, 0xB7, 0x4E), // yellow-green
            Color.FromRgb(0xD6, 0x74, 0x96), // rose
            Color.FromRgb(0x4E, 0xB7, 0xB7), // teal/cyan
            Color.FromRgb(0xC2, 0xAA, 0x56), // gold
            Color.FromRgb(0x7A, 0x8E, 0xC2), // indigo/slate blue
        };

        private static readonly SolidColorBrush[] PaletteBrushes;

        static RecentFileProvider()
        {
            PaletteBrushes = new SolidColorBrush[VsProjectColorPalette.Length];
            for (int i = 0; i < VsProjectColorPalette.Length; i++)
            {
                PaletteBrushes[i] = new SolidColorBrush(VsProjectColorPalette[i]);
                PaletteBrushes[i].Freeze();
            }
        }

        /// <summary>
        /// Builds a map of project name to color brush by enumerating projects in solution order.
        /// Colors are assigned sequentially from the VS palette, cycling as needed.
        /// </summary>
        private static async System.Threading.Tasks.Task<Dictionary<string, SolidColorBrush>> GetProjectColorMapAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var map = new Dictionary<string, SolidColorBrush>(System.StringComparer.OrdinalIgnoreCase);

            try
            {
                var solution = await VS.GetServiceAsync<SVsSolution, IVsSolution>();
                if (solution == null)
                    return map;

                System.Guid guid = System.Guid.Empty;
                int hr = solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, ref guid, out IEnumHierarchies enumHierarchies);
                if (hr != 0 || enumHierarchies == null)
                    return map;

                IVsHierarchy[] hierarchies = new IVsHierarchy[1];
                uint fetched;
                int colorIndex = 0;

                while (enumHierarchies.Next(1, hierarchies, out fetched) == 0 && fetched == 1)
                {
                    IVsHierarchy hierarchy = hierarchies[0];
                    if (hierarchy == null)
                        continue;

                    // Get the project name
                    hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out object nameObj);
                    string name = nameObj as string;

                    if (!string.IsNullOrEmpty(name) && !map.ContainsKey(name))
                    {
                        map[name] = PaletteBrushes[colorIndex % PaletteBrushes.Length];
                        colorIndex++;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
            }

            return map;
        }
    }
}
