using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    internal partial class Program
    {
        static void ParsingCSV(string path, string tableType = null)
        {
            StringBuilder stb = new StringBuilder();
            var sheetName = Path.GetFileName(path);
            var text = File.ReadAllText(path, Encoding.UTF8);

            string[] lines = text.Split("\r\n");

            Dictionary<int, string> columns = new Dictionary<int, string>();
            Dictionary<int, string> types = new Dictionary<int, string>();
            Dictionary<int, string> resultDatas = new Dictionary<int, string>();

            for (int i = 0; i < lines.Length; i++)
            {
                if (i == COLUMN_ROW)
                {
                    string[] datas = _DataRegex.Split(lines[i]);
                    for (int col = DATA_START_COL; col < datas.Length; col++)
                    {
                        if (string.IsNullOrWhiteSpace(datas[col]) == true)
                        {
                            continue;
                        }
                        columns.Add(col, datas[col].Trim().Replace("\"",""));
                    }
                    //Console.WriteLine(string.Join(",", columns.Values));
                    stb.AppendLine(string.Join(",", columns.Values));
                }
                else if (i == TYPE_ROW)
                {
                    string[] datas = _DataRegex.Split(lines[i]);
                    foreach (var col in columns)
                    {
                        types[col.Key] = datas[col.Key].Trim().Replace("\"", "");
                    }


                    //Console.WriteLine(string.Join(",", types.Values));
                    stb.AppendLine(string.Join(",", types.Values));
                }
                else if (i >= DATA_START_ROW)
                {

                    //Console.SetCursorPosition(0, Console.CursorTop - 1);
                    //Console.WriteLine($"[{sheetName}] Parsing Data {i}/{lines.Length}");

                    string line = lines[i];
                    if (line.Length == 0
                        || line.Trim().All(t => t == ',')) continue;

                    line = line.Replace("\n", "\\n");

                    string[] datas = _DataRegex.Split(line);

                    foreach (var col in columns)
                    {
                        resultDatas[col.Key] = datas[col.Key].Trim();
                    }

                    //Console.WriteLine(string.Join(",",resultDatas.Values));

                    if (resultDatas.Values.All(t => string.IsNullOrEmpty(t))) continue;

                    stb.AppendLine(string.Join(",", resultDatas.Values));
                }
            }


            if (string.IsNullOrEmpty(tableType) == false)
            {
                if (Directory.Exists($"{Program.EXCEL_FOLDER_DIR}/{tableType}/") == false)
                {
                    Directory.CreateDirectory($"{Program.EXCEL_FOLDER_DIR}/{tableType}/");
                }

                File.WriteAllText($"{Program.EXCEL_FOLDER_DIR}/{tableType}/{sheetName}", stb.ToString());
            }
            else
            {
                File.WriteAllText($"{Config.Default.EXTRACT_TARGET.FullName}/{sheetName}", stb.ToString());
                ///GenerateClass
                GenerateCS(sheetName, columns, types);
            }
        }

    }
}
