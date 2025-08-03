using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    internal static class FileExtension
    {
        public static string ProcessCreateFile(string dirPath, string fileName, string fileBody)
        {
            return SaveTextFileSafety($"{dirPath}/{fileName}", fileBody);
        }


        public static string SaveTextFileSafety(string path, string text)
        {
            var dir = Path.GetDirectoryName(path);

            if (dir != null && Directory.Exists(dir) == false)
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllText($"{path}", text, Encoding.UTF8);
            Console.WriteLine($"Save file {path}.");

            return path;
        }
    }
}
