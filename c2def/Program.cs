using G_PROJECT;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace c2def
{
    class Program
    {
        static void Main(string[] args)
        {
            var sourceDirectory = Environment.CurrentDirectory + @"\" + (args.Length != 0 ? args[0] : @"src");
            if (!Directory.Exists(sourceDirectory)) return;

            var allFilePaths = Directory.GetFiles(sourceDirectory, "*", SearchOption.AllDirectories);
            var targetFileExtensions = new List<string>() { "bas", "cls" };
            allFilePaths
                .Where(path => targetFileExtensions.Contains(Path.GetExtension(path)))
                .Where(path => GetEncoding(path) != Encoding.Default)
                .ToList().ForEach(path => Convert2DefaultEncoding(path));
        }

        private static void Convert2DefaultEncoding(string path)
        {
            var encoding = new TxtEnc().SetFromTextFile(path);
            var allText = File.ReadAllText(path, encoding);
            File.WriteAllText(path, allText, Encoding.Default);
        }

        private static Encoding GetEncoding(string path)
        {
            var te = new TxtEnc();
            return te.SetFromTextFile(path);
        }
    }
}
