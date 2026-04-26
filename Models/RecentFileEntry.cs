using System;

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

        public RecentFileEntry(string fullPath, string project, DateTime modified)
        {
            FullPath = fullPath;
            FileName = System.IO.Path.GetFileName(fullPath);
            Project = project;
            Modified = modified;
        }
    }
}
