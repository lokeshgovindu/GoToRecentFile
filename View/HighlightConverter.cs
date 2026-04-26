using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;

namespace GoToRecentFile.View
{
    /// <summary>
    /// Converts a file name and search terms into a TextBlock with highlighted matching portions.
    /// </summary>
    internal sealed class HighlightConverter : IMultiValueConverter
    {
        public Brush HighlightBrush { get; set; } = Brushes.Yellow;

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2)
                return new TextBlock();

            string text = values[0] as string ?? string.Empty;
            string[] searchWords = values[1] as string[] ?? Array.Empty<string>();

            var textBlock = new TextBlock();

            if (searchWords.Length == 0 || string.IsNullOrEmpty(text))
            {
                textBlock.Text = text;
                return textBlock;
            }

            // Find all match ranges
            var highlights = new List<Tuple<int, int>>();
            foreach (string word in searchWords)
            {
                if (string.IsNullOrEmpty(word))
                    continue;

                int index = 0;
                while ((index = text.IndexOf(word, index, StringComparison.OrdinalIgnoreCase)) >= 0)
                {
                    highlights.Add(Tuple.Create(index, index + word.Length));
                    index += word.Length;
                }
            }

            if (highlights.Count == 0)
            {
                textBlock.Text = text;
                return textBlock;
            }

            // Merge overlapping ranges
            var merged = MergeRanges(highlights);

            int pos = 0;
            foreach (var range in merged)
            {
                if (range.Item1 > pos)
                {
                    textBlock.Inlines.Add(new Run(text.Substring(pos, range.Item1 - pos)));
                }

                textBlock.Inlines.Add(new Run(text.Substring(range.Item1, range.Item2 - range.Item1))
                {
                    Background = HighlightBrush,
                    FontWeight = FontWeights.Bold
                });
                pos = range.Item2;
            }

            if (pos < text.Length)
            {
                textBlock.Inlines.Add(new Run(text.Substring(pos)));
            }

            return textBlock;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        private static List<Tuple<int, int>> MergeRanges(List<Tuple<int, int>> ranges)
        {
            var sorted = ranges.OrderBy(r => r.Item1).ThenBy(r => r.Item2).ToList();
            var result = new List<Tuple<int, int>> { sorted[0] };

            for (int i = 1; i < sorted.Count; i++)
            {
                var last = result[result.Count - 1];
                if (sorted[i].Item1 <= last.Item2)
                {
                    result[result.Count - 1] = Tuple.Create(last.Item1, Math.Max(last.Item2, sorted[i].Item2));
                }
                else
                {
                    result.Add(sorted[i]);
                }
            }

            return result;
        }
    }
}