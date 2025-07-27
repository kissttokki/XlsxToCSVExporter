using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    internal partial class Processing
    {
        public static void GenerateEnum(string path)
        {
            var sheetName = Path.GetFileName(path);
            var text = File.ReadAllText(path, Encoding.UTF8);

            var lines = text.Split("\r\n");


            CSVWriter clientCSV = new CSVWriter();
            CSVWriter serverCSV = new CSVWriter();

            Dictionary<int, CSVWriter[]> buildTargetWirters = new Dictionary<int, CSVWriter[]>();

            {
                string[] colRowDatas = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.EnumColumnNameRow));
                string[] buildTargetRows = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.EnumBuildTargetRow));


                for (int col = 0; col < buildTargetRows.Length; col++)
                {
                    if (string.IsNullOrWhiteSpace(colRowDatas[col]) == true || colRowDatas[col].StartsWith('#') == true)
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(buildTargetRows[0]) == true)
                    {
                        Console.BackgroundColor = ConsoleColor.Red;
                        Console.WriteLine($"[Warning] {sheetName} : 컬럼명{(colRowDatas[col])}의 빌드 타겟이 없습니다. 스킵됩니다.");
                        Console.BackgroundColor = default;
                        colRowDatas[col] = string.Empty;
                        continue;
                    }
                    {
                        var value = buildTargetRows[col].ToLower();
                        if (value == "both")
                        {
                            buildTargetWirters.Add(col, new CSVWriter[] { clientCSV, serverCSV });
                        }
                        else if (value == "client")
                        {
                            buildTargetWirters.Add(col, new CSVWriter[] { clientCSV });
                        }
                        else if (value == "server")
                        {
                            buildTargetWirters.Add(col, new CSVWriter[] { serverCSV });
                        }
                        else
                        {
                            Console.BackgroundColor = ConsoleColor.Red;
                            Console.WriteLine($"[Warning] {sheetName} : 컬럼명{(colRowDatas[col])}의 빌드 타겟 {buildTargetRows[col].ToLower()}을 알 수 없습니다. 스킵됩니다.");
                            Console.BackgroundColor = default;
                            Console.WriteLine();
                            continue;
                        }

                        foreach (var writer in buildTargetWirters[col])
                        {
                            writer.columns.Add(col, colRowDatas[col]);
                        }
                    }
                }


                for (int row = Config.Default.DataRow; row < lines.Count(); row++)
                {
                    string line = lines.ElementAt(row);
                    if (line.Length == 0 || line.Trim().All(t => t == ',')) continue;

                    line = line.Replace("\n", "\\n");

                    string[] datas = Config.Default.DataRegex.Split(line);

                    for (int dataCol = 0; dataCol < datas.Length; dataCol++)
                    {
                        if (buildTargetWirters.TryGetValue(dataCol, out CSVWriter[]? targetWriters) == true)
                        {
                            string cell = datas[dataCol];

                            foreach (var writer in targetWriters)
                            {
                                if (writer.resultDataLines.ContainsKey(row) == false)
                                {
                                    writer.resultDataLines[row] = new Dictionary<int, string>();
                                }
                                writer.resultDataLines[row][dataCol] = cell;
                            }
                        }
                    }

                    if (datas.All(t => string.IsNullOrWhiteSpace(t))) continue;
                }
            }

            foreach (var enumData in clientCSV.GetEnumCodes())
            {
                if (string.IsNullOrWhiteSpace(Config.Default.OutputClientCsharpScriptDir) == false)
                    FileExtension.SaveTextFileSafety($"{Config.Default.OutputClientCsharpScriptDir}/{enumData.name}.cs", enumData.code);
                else
                    FileExtension.SaveTextFileSafety($"output/client/cs/{enumData.name}.cs", enumData.code);
            }

            foreach (var enumData in serverCSV.GetEnumCodes())
            {
                if (string.IsNullOrWhiteSpace(Config.Default.OutputServerCsharpScriptDir) == false)
                    FileExtension.SaveTextFileSafety($"{Config.Default.OutputServerCsharpScriptDir}/{enumData.name}.cs", enumData.code);
                else
                    FileExtension.SaveTextFileSafety($"output/server/cs/{enumData.name}.cs", enumData.code);
            }
        }
    }
}
