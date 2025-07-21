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
    internal partial class Program
    {
        static void GenerateCS(string sheetName, Dictionary<int, string> cols, Dictionary<int, string> types)
        {
            string className = sheetName.Replace(".csv", "");
            string resultPath = $"{Config.Default.EXTRACT_CLASS_PATH}/Table{className}.cs";


            if (File.Exists(resultPath) == true)
            {
                //Console.WriteLine($"[{className}] 이미 Table{className}.cs 파일이 존재하여 cs파일을 갱신합니다.");

                ModifyCS(resultPath, className, cols, types);

                Console.WriteLine();
                return;
            }

            resultPath = $"{Config.Default.EXTRACT_CLASS_PATH}/AutoTable/Table{className}.cs";

            var template = File.ReadAllText($"{Environment.CurrentDirectory}/ClassTamplate.cs");



            var stb = new StringBuilder();
            var wrapperStb = new StringBuilder();

            var indexKey = 0;
            if (types.ContainsValue("index") == true)
            {
                indexKey = types.FirstOrDefault(t => t.Value == "index").Key;
            }
            else
            {
                indexKey = cols.Keys.First();
            }

            string indexTarget = cols[indexKey];
            string indexType = types[indexKey];


            foreach (KeyValuePair<int, string> item in cols)
            {
                if (item.Key == indexKey)
                {
                    if (types[indexKey].ToLower() == "index")
                        indexType = "int";
                    
                    stb.AppendLine($"\tpublic {indexType} {item.Value};");
                }
                else
                {
                    stb.AppendLine($"\tpublic {types[item.Key]} {item.Value};");
                }
                wrapperStb.AppendLine($"\t    data.{item.Value} = {item.Value};");
            }



            template = template.Replace("$CLASSNAME", className).Replace("$BODY", stb.ToString())
                .Replace("$WARPPERBODY", wrapperStb.ToString())
                .Replace("$INDEX_TYPE", indexType)
                .Replace("$INDEX_KEY", indexTarget)
                ;


            template = CSharpSyntaxTree.ParseText(template).GetRoot().NormalizeWhitespace().ToFullString();

            File.WriteAllText(resultPath, template);
            Console.WriteLine();
        }

        static void ModifyCS(string fileName, string className, Dictionary<int, string> cols, Dictionary<int, string> types)
        {
            var template = File.ReadAllText(fileName);
            CompilationUnitSyntax root = (CompilationUnitSyntax)CSharpSyntaxTree.ParseText(template).GetRoot();
            var classes = root.DescendantNodes().OfType<ClassDeclarationSyntax>();

            ClassDeclarationSyntax wrapperClass = null;
            ClassDeclarationSyntax dataClass = null;

            ///Class 로드
            foreach (ClassDeclarationSyntax c in classes)
            {
                if (c.Identifier.ToString() == $"{className}Wrapper")
                {
                    wrapperClass = c;
                }
                if (c.Identifier.ToString() == $"{className}")
                {
                    dataClass = c;
                }
            }


            if (dataClass == null)
            {
                Console.WriteLine($"[{className}] {className} 데이터클래스를 찾을 수 없어 수정할 수 없습니다.");

                Console.WriteLine("[클래스 리스트]");
                foreach (ClassDeclarationSyntax c in classes)
                {
                    Console.WriteLine($"\t{c.Identifier.ToString()}");
                }

                return;
            }
            if (wrapperClass == null)
            {
                Console.WriteLine($"[{className}] {className}Wrapper 래퍼클래스를 찾을 수 없어 수정할 수 없습니다.");

                Console.WriteLine("[클래스 리스트]");
                foreach (ClassDeclarationSyntax c in classes)
                {
                    Console.WriteLine($"\t{c.Identifier.ToString()}");
                }

                return;
            }



            ///기존 멤버라인 뒤에 추가하기 위해 멤버들 가져옴
            List<MemberDeclarationSyntax> wrapperMembers = wrapperClass.Members.ToList();
            List<string> wrapperMemberNames = wrapperMembers
                .OfType<FieldDeclarationSyntax>()
                .Select(t => t.Declaration.Variables.First().Identifier.ToString())
                .ToList();



            ///index의 키값을 알아내기 위해..
            ///참조하는 인터페이스 중 ITableData<T> 찾기
            var interfaceType = dataClass.BaseList?.Types
                .FirstOrDefault(t => t.Type is GenericNameSyntax &&
                                     ((GenericNameSyntax)t.Type).Identifier.Text == "ITableData");
            string indexKey = null;
            if (interfaceType != null)
            {
                // 제네릭 타입 인수 추출
                var genericTypeArgument = ((GenericNameSyntax)interfaceType.Type).TypeArgumentList.Arguments.First();
                indexKey = genericTypeArgument.ToString();
            }
            else
            {
                Console.WriteLine("DataClass가 ITableData<T> 인터페이스를 참조하지 않습니다.");
                return;
            }


            foreach (var kv in cols)
            {
                if (wrapperMemberNames.Contains(kv.Value) == true)
                {
                    var member = wrapperMembers.OfType<FieldDeclarationSyntax>()
                                   .FirstOrDefault(t => t.Declaration.Variables.First().Identifier.ToString() == kv.Value);

                    if (member != null)
                    {
                        string currentType = member.Declaration.Type.ToString();
                        string targetType = types[kv.Key].ToLower() == "index" ? indexKey : types[kv.Key];

                        if (currentType != targetType)
                        {
                            Console.WriteLine($"[{className}] 타입이 다릅니다!! Modifying type of {kv.Value} from {currentType} to {targetType}");

                            var newFieldDeclaration = SyntaxFactory.FieldDeclaration(
                                SyntaxFactory.VariableDeclaration(
                                    SyntaxFactory.ParseTypeName(targetType),
                                    SyntaxFactory.SeparatedList(new[] { SyntaxFactory.VariableDeclarator(SyntaxFactory.Identifier(kv.Value)) })
                                )
                            ).AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));

                            wrapperMembers[wrapperMembers.IndexOf(member)] = newFieldDeclaration;
                        }
                    }


                    continue;
                }

                string typeName = types[kv.Key];

                if (typeName == "index")
                {
                    typeName = indexKey;
                    Console.WriteLine($"[{className}] 신규 Index 데이터 감지.. 추가 데이터 {kv.Value}, 타입 : {typeName}");
                }
                else
                    Console.WriteLine($"[{className}] 신규 데이터 감지.. 추가 데이터 {kv.Value}, 타입 : {typeName}");


                // 멤버 선언 생성
                var fieldDeclaration = SyntaxFactory.FieldDeclaration(
                    SyntaxFactory.VariableDeclaration(
                        SyntaxFactory.ParseTypeName(typeName),
                        SyntaxFactory.SeparatedList(new[] { SyntaxFactory.VariableDeclarator(SyntaxFactory.Identifier(kv.Value)) })
                    )
                )
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));

                // 멤버 추가
                wrapperMembers.Insert(wrapperMembers.FindLastIndex(m => m is FieldDeclarationSyntax) + 1, fieldDeclaration);
            }


            root = root.ReplaceNode(wrapperClass, wrapperClass.WithMembers(SyntaxFactory.List<MemberDeclarationSyntax>(wrapperMembers)));


            File.WriteAllText(fileName, root.NormalizeWhitespace().ToFullString());
        }

    }
}
