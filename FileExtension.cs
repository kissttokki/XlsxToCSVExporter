using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    internal static class FileExtension
    {
        public static void ProcessCreateFile(string dirPath, string fileName, string fileBody, string addKey = null)
        {
            if (string.IsNullOrWhiteSpace(dirPath) == false)
            {
                SaveTextFileSafety($"{dirPath}/{fileName}", fileBody);
            }
            else
            {
                SaveTextFileSafety($"output{addKey}/{fileName}", fileBody);
            }
        }


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
