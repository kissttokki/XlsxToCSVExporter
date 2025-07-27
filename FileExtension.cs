using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    internal static class FileExtension
    {
        public static void SaveTextFileSafety(string path, string text)
        {
            var dir = Path.GetDirectoryName(path);

            if (dir != null && Directory.Exists(dir) == false)
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllText($"{path}", text, Encoding.UTF8);
            Console.WriteLine($"Save file {path}.");
        }
    }
}
