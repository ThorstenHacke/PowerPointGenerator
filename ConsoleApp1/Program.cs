using PowerPointGenerator;
using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please provide the path to the Template Presentation:");
            string pptxPath = Console.ReadLine();
            Console.WriteLine("Please provide the path to the Data-CSV-File:");
            string csvPath = Console.ReadLine();
            //Generator.Validate();
            var generator = new Generator(pptxPath, csvPath);
            generator.Convert();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Generation done. Where do you want to save the presentation:");
            Console.ForegroundColor = ConsoleColor.White;

            string savePath = Console.ReadLine();
            generator.Save(savePath);
            Console.WriteLine("Presentation saved. Press any key to close");
            Console.ReadKey();
        }
    }
}
