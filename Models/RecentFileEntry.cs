using System;
using System.Windows.Media;

namespace GoToRecentFile.Models
{
    /// <summary>
    /// Represents a recently opened file with metadata for display.
    /// </summary>
    internal sealed class RecentFileEntry
    {
        public string FileName { get; }
        public string Project { get; }
        public string FullPath { get; }
        public DateTime Modified { get; }

        /// <summary>
        /// The project color as assigned by VS tab colorization (or Transparent if unavailable).
        /// </summary>
        public SolidColorBrush ProjectColorBrush { get; }

        public RecentFileEntry(string fullPath, string project, DateTime modified, SolidColorBrush projectColorBrush = null)
        {
            FullPath = fullPath;
            FileName = System.IO.Path.GetFileName(fullPath);
            Project = project;
            Modified = modified;
            ProjectColorBrush = projectColorBrush ?? Brushes.Transparent;
        }
    }
}
