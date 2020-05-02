using System;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    public class AutoCompletionHandler : IAutoCompleteHandler
    {
        private const char Separator = '\\';

        // characters to start completion from
        public char[] Separators { get; set; } = new char[] {  };

        // text - The current text entered in the console
        // index - The index of the terminal cursor within {text}
        public string[] GetSuggestions(string text, int index)
        {
            if (text.Length == 0)
            {
                return Environment.GetLogicalDrives();
            }
            if (!text.Contains(Separator))
            {
                var drives = Environment.GetLogicalDrives().Where(ld => ld.StartsWith(text, StringComparison.OrdinalIgnoreCase));
                return drives.ToArray();
            }
            var parts = text.Split(Separator);
            var knownPart = text.Substring(0, text.LastIndexOf(Separator))+ Separator;
            var entries = Directory.EnumerateFileSystemEntries(knownPart);
            var newPart = parts.Last();
            var found = entries.Where(e => IsPathMatch(e, knownPart, newPart));
            return found.ToArray();
        }

        private bool IsPathMatch(string path, string knownPart, string newPart) {
            var cleanedPath = path.Replace(knownPart, string.Empty);
            return cleanedPath.StartsWith(newPart);
        }
    }

}
