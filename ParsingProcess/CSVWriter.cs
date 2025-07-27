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
               SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System.Collections.Generic")),
               SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("Newtonsoft.Json"))
               );


            if (string.IsNullOrWhiteSpace(Config.Default.ResultNameSpace) == false)
            {
                compilationUnit = compilationUnit.AddMembers(SyntaxFactory.NamespaceDeclaration(SyntaxFactory.ParseName($"{Config.Default.ResultNameSpace}")));
            }

            var attr = SyntaxFactory.AttributeList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Attribute(SyntaxFactory.IdentifierName("JsonProperty"))));


            List<PropertyDeclarationSyntax> members = new List<PropertyDeclarationSyntax>();

            foreach (var col in columns)
            {
                members.Add(SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName(dataTypes[col.Key]), col.Value)
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .AddAccessorListAccessors(
                    SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                    .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                    SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                    .WithModifiers(new SyntaxTokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)))
                    .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)))
                .AddAttributeLists(attr));
            }


            var classDeclaration = SyntaxFactory.ClassDeclaration(className)
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .AddMembers(items : members.ToArray());


            compilationUnit = compilationUnit.AddMembers(classDeclaration);

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
            var compilationUnit = SyntaxFactory.CompilationUnit()
                .AddUsings(SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System")))
                .AddMembers(SyntaxFactory.EnumDeclaration(enumName).AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword)).WithMembers(SyntaxFactory.SeparatedList(members)));


            if (string.IsNullOrWhiteSpace(Config.Default.ResultNameSpace) == false)
            {
                compilationUnit = compilationUnit.AddMembers(SyntaxFactory.NamespaceDeclaration(SyntaxFactory.ParseName($"{Config.Default.ResultNameSpace}")));
            }

            return compilationUnit.NormalizeWhitespace().ToString();
        }
    }

}
