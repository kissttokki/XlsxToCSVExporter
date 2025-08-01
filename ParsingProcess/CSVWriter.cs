using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace TableExporter
{
    class CSVWriter
    {
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

        public string GetCSV()
        {
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

        public string GetClassCode(string className)
        {
            var compilationUnit = SyntaxFactory.CompilationUnit()
               .AddUsings(
                   SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System")),
                   SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System.Collections.Generic"))
               );

            //var attr = SyntaxFactory.AttributeList(
            //    SyntaxFactory.SingletonSeparatedList(
            //        SyntaxFactory.Attribute(SyntaxFactory.IdentifierName("JsonProperty"))
            //    )
            //);


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

            MethodDeclarationSyntax method = null;

            if (indexCol.Value == null)
            {
                members.Add(SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName("int"), "AutoGenrate_INDEX")
                    .AddModifiers(SyntaxFactory.Token(SyntaxKind.ProtectedKeyword))
                    .AddAccessorListAccessors(
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                            .WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)))
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken))
                    ));


                method = SyntaxFactory.MethodDeclaration(SyntaxFactory.ParseTypeName(indexType), "GetKey")
                  .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword), SyntaxFactory.Token(SyntaxKind.OverrideKeyword))
                  .WithExpressionBody(SyntaxFactory.ArrowExpressionClause(SyntaxFactory.IdentifierName("AutoGenrate_INDEX")))
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
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .AddBaseListTypes(SyntaxFactory.SimpleBaseType(SyntaxFactory.ParseTypeName($"BaseTableData<{indexType}>")))
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
