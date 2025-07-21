using Microsoft.CodeAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    internal partial class Program
    {

        static void ParsingEnum(string path)
        {
            var sheetName = Path.GetFileName(path);
            var text = File.ReadAllText(path, Encoding.UTF8);

            var lines = text.Split("\r\n");

            string csName = sheetName.Replace(".csv", ".cs").Replace("enum_", "");

            var template = File.ReadAllText($"{Environment.CurrentDirectory}/EnumTamplate.cs");

            Dictionary<int, string> columns = new Dictionary<int, string>();
            Dictionary<int, string> types = new Dictionary<int, string>();
            Dictionary<string, string> resultDatas = new Dictionary<string, string>();
            Dictionary<string, string> annotations = new Dictionary<string, string>();

            for (int i = 0; i < lines.Length; i++)
            {
                if (i >= DATA_START_ROW)
                {
                    string line = lines[i];
                    if (line.Length == 0
                        || line.Trim().All(t => t == ',')) continue;

                    try
                    {
                        string[] datas = _DataRegex.Split(line).Where(t => string.IsNullOrEmpty(t) == false).Take(3).ToArray();


                        string enumValue;
                        string enumKey;

                        if (int.TryParse(datas[1], out _) == true)
                        {
                            enumKey = datas[0];
                            enumValue = datas[1];
                        }
                        else
                        {
                            enumKey = datas[1];
                            enumValue = datas[0];
                        }


                        if(datas.Length > 2 && string.IsNullOrWhiteSpace(datas[2]) == false)
                            annotations[enumKey] = datas[2];

                        resultDatas[enumKey] = enumValue;

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Enum Parsing Error : {ex}");
                    }
                }
            }

            StringBuilder body = new StringBuilder();

            foreach (var item in resultDatas)
            {
                if (annotations.ContainsKey(item.Key) == true)
                { 
                    body.AppendLine("/// <summary>");
                    body.AppendLine($"/// {annotations[item.Key]}");
                    body.AppendLine("/// </summary>");
                }

                if (string.IsNullOrEmpty(item.Value) == false)
                    body.AppendLine($"\t\t{item.Key} = {item.Value},");
                else
                    body.AppendLine($"\t\t{item.Key},");
            }

            var enumName = csName.Replace(".cs", "");

            template = template.Replace("$TITLE", enumName).Replace("$BODY", body.ToString());
            template = Microsoft.CodeAnalysis.CSharp.CSharpSyntaxTree.ParseText(template).GetRoot().NormalizeWhitespace().ToFullString();
            File.WriteAllText($"{Config.Default.EXTRACT_ENUM_TARGET.FullName}/{csName}", template);
        }
    }
}
