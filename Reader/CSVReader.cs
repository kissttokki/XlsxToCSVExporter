using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Security.Cryptography;

namespace TableExporter
{
    class CSVReader
    {
        private static readonly HashSet<string> s_numericTypes 
            = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "int", "long", "short", "float", "double", "uint", "ulong", "ushort", "Int32", "Int64", "UInt32", "UInt64" };

        public enum CSVWriteType
        {
            None = 0,
            Server = 1 << 0,
            Client = 1 << 1,
        }

        /// <summary>
        /// key - column
        /// </summary>
        public Dictionary<int, string> columns = new Dictionary<int, string>();
        /// <summary>
        /// key - column
        /// </summary>
        public Dictionary<int, string> dataTypes = new Dictionary<int, string>();
        /// <summary>
        /// key - row
        /// </summary>
        public Dictionary<int, Dictionary<int, string>> resultDataLines = new();

        private string _csvPath;
        private CSVWriteType _buildType;

        public CSVReader(string csvPath, params CSVWriteType[] type)
        {
            _csvPath = csvPath;

            _buildType = CSVWriteType.None;
            foreach (var item in type)
            {
                _buildType |= item;
            }
        }

        private void LoadData()
        {
            if (columns.Count != 0)
                return;

            string sheetName = Path.GetFileName(_csvPath);
            string text = File.ReadAllText(_csvPath, Encoding.UTF8);
            string[] lines = text.Split("\r\n");

            Dictionary<int, bool> IsTargetDict = new Dictionary<int, bool>();

            string[] colRowDatas = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.ColumnNameRow));
            string[] buildTargetRows = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.BuildTargetRow));
            string[] dataTypeRowData = Config.Default.DataRegex.Split(lines.ElementAt(Config.Default.DataTypeRow));

            for (int col = 0; col < buildTargetRows.Length; col++)
            {
                if (string.IsNullOrWhiteSpace(colRowDatas[col]) == true || colRowDatas[col].StartsWith('#') == true)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(buildTargetRows[col]) == true)
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


                var value = buildTargetRows[col].ToLower();
                if (value == "both")
                {
                    IsTargetDict.Add(col, true);
                }
                else if (value == "client" && ((_buildType & CSVWriteType.Client) > 0))
                {
                    IsTargetDict.Add(col, true);
                }
                else if (value == "server" && ((_buildType & CSVWriteType.Server) > 0))
                {
                    IsTargetDict.Add(col, true);
                }
                else
                {
                    //Console.BackgroundColor = ConsoleColor.Red;
                    //Console.WriteLine($"[Warning][{_buildType}] {sheetName} : 컬럼명{(colRowDatas[col])}의 빌드 타겟 {buildTargetRows[col].ToLower()}을 알 수 없습니다. 스킵됩니다.");
                    //Console.BackgroundColor = default;
                    continue;
                }

                if (IsTargetDict[col] == true)
                {
                    columns.Add(col, colRowDatas[col]);
                    dataTypes.Add(col, dataTypeRowData[col].Replace("\"", ""));
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
                    if (IsTargetDict.TryGetValue(dataCol, out bool isTarget) == true)
                    {
                        string cell = datas[dataCol];
                        if (s_numericTypes.Contains(dataTypeRowData[dataCol]) == true)
                        {
                            cell = cell.Replace(",", "");
                        }

                        if (resultDataLines.ContainsKey(row) == false)
                        {
                            resultDataLines[row] = new Dictionary<int, string>();
                        }
                        resultDataLines[row][dataCol] = cell;
                    }
                }

                if (datas.All(t => string.IsNullOrWhiteSpace(t))) continue;
            }
        }

        private void LoadEnum()
        {
            var sheetName = Path.GetFileName(_csvPath);
            var text = File.ReadAllText(_csvPath, Encoding.UTF8);

            var lines = text.Split("\r\n");
            Dictionary<int, bool> IsTargetDict = new Dictionary<int, bool>();

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
                        IsTargetDict.Add(col, true);
                    }
                    else if (value == "client" && ((_buildType & CSVWriteType.Client) > 0))
                    {
                        IsTargetDict.Add(col, true);
                    }
                    else if (value == "server" && ((_buildType & CSVWriteType.Server) > 0))
                    {
                        IsTargetDict.Add(col, true);
                    }
                    else
                    {
                        Console.WriteLine();
                        continue;
                    }

                    if (IsTargetDict[col] == true)
                    {
                        columns.Add(col, colRowDatas[col]);
                    }
                }
            }


            for (int row = Config.Default.EnumDataRow; row < lines.Count(); row++)
            {
                string line = lines.ElementAt(row);
                if (line.Length == 0 || line.Trim().All(t => t == ',')) continue;

                line = line.Replace("\n", "\\n");

                string[] datas = Config.Default.DataRegex.Split(line);

                for (int dataCol = 0; dataCol < datas.Length; dataCol++)
                {
                    if (IsTargetDict.TryGetValue(dataCol, out var result) == true)
                    {
                        string cell = datas[dataCol];
                        if (resultDataLines.ContainsKey(row) == false)
                        {
                            resultDataLines[row] = new Dictionary<int, string>();
                        }
                        resultDataLines[row][dataCol] = cell;
                    }
                }

                if (datas.All(t => string.IsNullOrWhiteSpace(t))) continue;
            }
        }


        public string GetCSV()
        {
            LoadData();
            StringBuilder stb = new StringBuilder();

            stb.AppendLine(string.Join(',', columns.Values));

            if (dataTypes.Count > 0)
            {
                stb.AppendLine(string.Join(',', dataTypes.Values));
            }

            foreach (var rowData in resultDataLines)
            {
                stb.AppendLine(string.Join(',', rowData.Value.Values));
            }

            return stb.ToString();
        }

        public string GetClassCode()
        {
            LoadData();

            string className = Path.GetFileName(_csvPath).Replace(".csv", "");

            var compilationUnit = SyntaxFactory.CompilationUnit()
               .AddUsings(
                   SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System")),
                   SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System.Collections.Generic"))
               );



            var indexCol = columns.FirstOrDefault(col => dataTypes[col.Key].Equals("Index", StringComparison.OrdinalIgnoreCase));
            string indexType = indexCol.Value == null || (int.TryParse(resultDataLines.First().Value[indexCol.Key], out _) == true) ? "int" : "string";

            ///멤버 생성단
            var members = columns.Select(col =>
            {
                PropertyDeclarationSyntax prop = SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName(col.Key == indexCol.Key ? indexType : dataTypes[col.Key]), col.Value)
                    .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                    .AddAccessorListAccessors(
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                            .WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)))
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken))
                    );
                return prop;
            }).ToList();

            List<FieldDeclarationSyntax> fieldMembers = new List<FieldDeclarationSyntax>();

            MethodDeclarationSyntax method = null;

            if (indexCol.Value == null)
            {
                // public int Index_AutoIncremented { get; private set; }
                members.Add(
                    SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName("int"), "Index_AutoIncremented")
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                        .AddAccessorListAccessors(
                            SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                                .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                            SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                                .WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)))
                                .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken))
                        )
                );

                method = SyntaxFactory.MethodDeclaration(SyntaxFactory.ParseTypeName(indexType), "GetKey")
                  .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword), SyntaxFactory.Token(SyntaxKind.OverrideKeyword))
                  .WithExpressionBody(SyntaxFactory.ArrowExpressionClause(SyntaxFactory.IdentifierName("Index_AutoIncremented")))
                  .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken));
            }
            else
            {
                method = SyntaxFactory.MethodDeclaration(SyntaxFactory.ParseTypeName(indexType), "GetKey")
                    .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword), SyntaxFactory.Token(SyntaxKind.OverrideKeyword))
                    .WithExpressionBody(SyntaxFactory.ArrowExpressionClause(SyntaxFactory.IdentifierName(indexCol.Value)))
                    .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken));
            }


            ///데이터 클래스
            ClassDeclarationSyntax dataClassDeclaration = SyntaxFactory.ClassDeclaration($"{className}Data")
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword), SyntaxFactory.Token(SyntaxKind.PartialKeyword))
                .AddBaseListTypes(SyntaxFactory.SimpleBaseType(SyntaxFactory.ParseTypeName($"BaseTableData<{indexType}>")))
                .AddMembers(fieldMembers.ToArray())
                .AddMembers(members.ToArray())
                .AddMembers(method);


            ClassDeclarationSyntax tableClassDeclaration = SyntaxFactory.ClassDeclaration(className)
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword), SyntaxFactory.Token(SyntaxKind.PartialKeyword))
                .AddBaseListTypes(SyntaxFactory.SimpleBaseType(SyntaxFactory.ParseTypeName($"BaseTable<{className}Data, {indexType}>")));



            if (string.IsNullOrWhiteSpace(Config.Default.ResultNameSpace) == false)
            {
                compilationUnit = compilationUnit.AddMembers(
                    SyntaxFactory.NamespaceDeclaration(
                    SyntaxFactory.ParseName(Config.Default.ResultNameSpace)
                ).AddMembers(dataClassDeclaration, tableClassDeclaration));
            }
            else
            {
                compilationUnit = compilationUnit.AddMembers(dataClassDeclaration, tableClassDeclaration);
            }


            return compilationUnit.NormalizeWhitespace().ToFullString();
        }

        public IEnumerable<(string name,string code)> GetEnumCodes()
        {
            LoadEnum();
            List<(string name, string code)> results = new ();

            List<string> values = new List<string>();
            foreach (var col in columns)
            {
                foreach (var row in resultDataLines.Values)
                {

                    if (row.TryGetValue(col.Key, out var cellData) == true && string.IsNullOrWhiteSpace(cellData) == false)
                    {
                        values.Add(cellData);
                    }
                }

                results.Add((col.Value, GetEnumCSharpScript(col.Value, values)));

                values.Clear();
            }

            return results;
        }


        private string GetEnumCSharpScript(string enumName, IEnumerable<string> values)
        {
            var loanPurpose = values;
            var members = loanPurpose.Select(name =>
                SyntaxFactory.EnumMemberDeclaration(name)
            );

            var enumDecl = SyntaxFactory.EnumDeclaration(enumName)
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .WithMembers(SyntaxFactory.SeparatedList(members));

            var compilationUnit = SyntaxFactory.CompilationUnit()
                .AddUsings(SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System")));

            if (string.IsNullOrWhiteSpace(Config.Default.ResultNameSpace) == false)
            {
                var namespaceDecl = SyntaxFactory.NamespaceDeclaration(
                    SyntaxFactory.ParseName(Config.Default.ResultNameSpace))
                    .AddMembers(enumDecl);

                compilationUnit = compilationUnit.AddMembers(namespaceDecl);
            }
            else
            {
                compilationUnit = compilationUnit.AddMembers(enumDecl);
            }

            return compilationUnit.NormalizeWhitespace().ToString();
        }
    }

}
