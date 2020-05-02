using PowerPointGenerator;
using System;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadLine.AutoCompletionHandler = new AutoCompletionHandler();
            Console.WriteLine("Please provide the path to the Template Presentation:");
            string pptxPath = ReadLine.Read();
            Console.WriteLine("Please provide the path to the Data-CSV-File:");
            string csvPath = ReadLine.Read();
            //Generator.Validate();
            var generator = new Generator(pptxPath, csvPath);
            generator.Convert();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Generation done. Where do you want to save the presentation:");
            Console.ForegroundColor = ConsoleColor.White;

            string savePath = ReadLine.Read();
            generator.SaveInDestination(savePath);
            Console.WriteLine("Presentation saved. Press any key to close");
            Console.ReadKey();
        }
    }

    class AutoCompletionHandler : IAutoCompleteHandler
    {
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
            var parts = text.Split('\\');
            var drives = Environment.GetLogicalDrives().Where(ld => ld.StartsWith(parts[0], StringComparison.OrdinalIgnoreCase));
            if (!text.Contains('\\'))
            {
                return drives.ToArray();
            }
            var knownPart = text.Substring(0, text.LastIndexOf('\\'))+"\\";
            var entries = Directory.EnumerateFileSystemEntries(knownPart);
            var found = entries.Where(e => e.Replace(knownPart,"").StartsWith(parts[parts.Length - 1]));
            return found.ToArray();
        }
    }

}
