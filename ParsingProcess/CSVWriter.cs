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

            ///멤버 생성단
            var members = columns.Select(col =>
            {
                var isIndexType = dataTypes[col.Key].Equals("Index", StringComparison.OrdinalIgnoreCase);
                var isIndexName = col.Value.Equals("Index", StringComparison.OrdinalIgnoreCase);

                var typeName = isIndexType ? "int" : dataTypes[col.Key];
                var prop = SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName(typeName), col.Value)
                    .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));

                /// Index 타입인데 컬럼명이 Index가 아닌 경우
                if (isIndexType && !isIndexName)
                {
                    prop = prop.AddAccessorListAccessors(
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                            .WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)))
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken))
                    );

                    /// 추가로 override Index 프로퍼티 생성해서 연결
                    var indexProp = SyntaxFactory.PropertyDeclaration(SyntaxFactory.ParseTypeName("int"), "Index")
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword), SyntaxFactory.Token(SyntaxKind.OverrideKeyword))
                        .AddAccessorListAccessors(
                            SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                                .WithExpressionBody(
                                    SyntaxFactory.ArrowExpressionClause(
                                        SyntaxFactory.IdentifierName(col.Value)
                                    )
                                )
                                .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                            SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                                .WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.ProtectedKeyword)))
                                .WithExpressionBody(
                                    SyntaxFactory.ArrowExpressionClause(
                                        SyntaxFactory.AssignmentExpression(
                                            SyntaxKind.SimpleAssignmentExpression,
                                            SyntaxFactory.IdentifierName(col.Value),
                                            SyntaxFactory.IdentifierName("value")
                                        )
                                    )
                                )
                                .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken))
                        );

                    return new[] { prop, indexProp };
                }
                else
                {
                    /// Index 컬럼명이면 override Index로 생성
                    if (isIndexType && isIndexName)
                    {
                        prop = prop.AddModifiers(SyntaxFactory.Token(SyntaxKind.OverrideKeyword));
                    }

                    prop = prop.AddAccessorListAccessors(
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken)),
                        SyntaxFactory.AccessorDeclaration(SyntaxKind.SetAccessorDeclaration)
                            .WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PrivateKeyword)))
                            .WithSemicolonToken(SyntaxFactory.Token(SyntaxKind.SemicolonToken))
                    );

                    return new[] { prop };
                }
            }).SelectMany(p => p).ToArray();


            var classDeclaration = SyntaxFactory.ClassDeclaration(className)
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .AddBaseListTypes(SyntaxFactory.SimpleBaseType(SyntaxFactory.ParseTypeName($"BaseTable<{className}>")))
                .AddMembers(members);

            MemberDeclarationSyntax finalMember = classDeclaration;

            if (!string.IsNullOrWhiteSpace(Config.Default.ResultNameSpace))
            {
                finalMember = SyntaxFactory.NamespaceDeclaration(
                    SyntaxFactory.ParseName(Config.Default.ResultNameSpace)
                ).AddMembers(classDeclaration);
            }

            compilationUnit = compilationUnit.AddMembers(finalMember);

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
