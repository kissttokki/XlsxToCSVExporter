using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{

    internal partial class Processing
    {
        public static void ParsingCSV(string path)
        {
            var numericTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "int", "long", "short", "float", "double", "uint", "ulong", "ushort", "Int32", "Int64", "UInt32", "UInt64" };

            CSVWriter clientCSV = new CSVWriter();
            CSVWriter serverCSV = new CSVWriter();

            string sheetName = Path.GetFileName(path);
            string text = File.ReadAllText(path, Encoding.UTF8);
            string[] lines = text.Split("\r\n");

            Dictionary<int, CSVWriter[]> buildTargetWirters = new Dictionary<int, CSVWriter[]>();

            string[] colRowDatas = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.ColumnNameRow));
            string[] buildTargetRows = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.BuildTargetRow));
            string[] dataTypeRowData = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.DataTypeRow));

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

                if (string.IsNullOrWhiteSpace(dataTypeRowData[col]) == true)
                {
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[Warning] {sheetName} : 컬럼명{(colRowDatas[col])}의 데이터 타입이 없습니다. 스킵됩니다.");
                    Console.BackgroundColor = default;
                    continue;
                }
                else
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
                        writer.dataTypes.Add(col, dataTypeRowData[col]);
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
                        if (numericTypes.Contains(dataTypeRowData[dataCol]) == true)
                        {
                            cell = cell.Replace(",", "");
                        }

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


            FileExtension.ProcessCreateFile(Config.Default.OutputClientCSVDir, sheetName, clientCSV.GetCSV(), "/client");
            FileExtension.ProcessCreateFile(Config.Default.OutputServerCSVDir, sheetName, serverCSV.GetCSV(), "/server");

            string className = sheetName.Replace(".csv", "");
            ///CSharpScript
            FileExtension.ProcessCreateFile(Config.Default.OutputClientCsharpScriptDir, $"{className}.cs", clientCSV.GetClassCode(className), "/client");
            FileExtension.ProcessCreateFile(Config.Default.OutputServerCsharpScriptDir, $"{className}.cs", serverCSV.GetClassCode(className), "/server");
        }
    }
}
