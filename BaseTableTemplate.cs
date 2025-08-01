using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;


public abstract class BaseTableData<K>
{
    public abstract K GetKey();
}

public abstract class BaseTable<T, K> where T : BaseTableData<K>, new()
{
    protected static readonly Dictionary<K, T> s_dict = new Dictionary<K, T>();

    public static IReadOnlyCollection<T> Datas => s_dict.Values;

    public static T GetValue(K index)
    {
        if (s_dict.TryGetValue(index, out T result) == true)
        {
            return result;
        }
        else
        {
#if UNITY_EDITOR
            UnityEngine.Debug.LogError($"[{typeof(T).Name}] not found index : {index}");
#else
            Console.WriteLine($"[{typeof(T).Name}] not found index : {index}");
#endif
        }

        return result;
    }


    public static void LoadCsv(string csvData)
    {
        s_dict.Clear();

        MemberInfo[] cachedMembers = null;
        MemberSetterDelegate[] cachedSetters = null;

        var currentLineBuilder = new StringBuilder();
        var fieldBuilder = new StringBuilder();
        var dataType = typeof(T);
        var colCount = 0;

        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData ?? ""));
        using var reader = new StreamReader(ms);

        string[] headerCells = null;
        bool isInQuotedField = false;

        while (!reader.EndOfStream)
        {
            string line = reader.ReadLine();
            if (string.IsNullOrEmpty(line)) continue;

            // 첫 줄: 헤더 처리
            if (headerCells == null)
            {
                headerCells = line.Split(',');
                int colLen = headerCells.Length;
                cachedMembers = new MemberInfo[colLen];
                cachedSetters = new MemberSetterDelegate[colLen];

                for (int col = 0; col < colLen; col++)
                {
                    string header = headerCells[col].Trim();
                    var member = dataType.GetField(header, BindingFlags.Public | BindingFlags.Instance)
                               ?? (MemberInfo)dataType.GetProperty(header, BindingFlags.Public | BindingFlags.Instance);

                    if (member != null)
                    {
                        cachedMembers[col] = member;
                        cachedSetters[col] = CreateMemberSetter(member);
                    }
                }

                // 타입 라인 스킵
                reader.ReadLine();
                continue;
            }

            if (isInQuotedField)
            {
                currentLineBuilder.AppendLine(line);
                if (line.Contains("\""))
                {
                    isInQuotedField = false;
                    ProcessLine(currentLineBuilder.ToString().AsSpan());
                    currentLineBuilder.Clear();
                }
            }
            else
            {
                int quoteCount = CountQuotes(line);
                if (quoteCount % 2 != 0)
                {
                    isInQuotedField = true;
                    currentLineBuilder.AppendLine(line);
                }
                else
                {
                    ProcessLine(line.AsSpan());
                    colCount++;
                }
            }
        }

        if (currentLineBuilder.Length > 0)
        {
            ProcessLine(currentLineBuilder.ToString().AsSpan());
            colCount++;
        }

        void ProcessLine(ReadOnlySpan<char> line)
        {
            int index = 0;
            int columnIndex = 0;
            T data = new T();

            while (index < line.Length)
            {
                if (line[index] == '"')
                {
                    index++;
                    fieldBuilder.Clear();

                    while (index < line.Length)
                    {
                        int quoteIndex = line.Slice(index).IndexOf('"');
                        if (quoteIndex == -1)
                        {
                            fieldBuilder.Append(line.Slice(index).ToString());
                            index = line.Length;
                            break;
                        }

                        fieldBuilder.Append(line.Slice(index, quoteIndex).ToString());
                        int quotePos = index + quoteIndex;

                        if (quotePos + 1 < line.Length && line[quotePos + 1] == '"')
                        {
                            fieldBuilder.Append('"');
                            index = quotePos + 2;
                        }
                        else
                        {
                            index = quotePos + 1;
                            break;
                        }
                    }

                    while (index < line.Length && char.IsWhiteSpace(line[index])) index++;
                    if (index < line.Length && line[index] == ',') index++;

                    string fieldValue = fieldBuilder.ToString();
                    if (!string.IsNullOrEmpty(fieldValue))
                    {
                        var memberType = GetMemberType(cachedMembers[columnIndex]);
                        var converted = ConvertType(memberType, fieldValue);
                        cachedSetters[columnIndex]?.Invoke(data, converted);
                    }

                    columnIndex++;
                }
                else
                {
                    int commaIndex = line.Slice(index).IndexOf(',');
                    ReadOnlySpan<char> token;

                    if (commaIndex == -1)
                    {
                        token = line.Slice(index).Trim();
                        index = line.Length;
                    }
                    else
                    {
                        token = line.Slice(index, commaIndex).Trim();
                        index += commaIndex + 1;
                    }

                    if (token.Length > 0)
                    {
                        var memberType = GetMemberType(cachedMembers[columnIndex]);
                        var converted = ConvertType(memberType, token.ToString());
                        cachedSetters[columnIndex]?.Invoke(data, converted);
                    }

                    columnIndex++;
                }
            }

            s_dict.Add(data.GetKey(), data);
        }

        int CountQuotes(string line)
        {
            int count = 0;
            foreach (var c in line) if (c == '"') count++;
            return count;
        }

        Type GetMemberType(MemberInfo member) =>
            member is FieldInfo fi ? fi.FieldType :
            member is PropertyInfo pi ? pi.PropertyType :
            throw new NotSupportedException();

        object ConvertType(Type type, string value)
        {
            if (type.IsEnum) return Enum.Parse(type, value);

            var nullable = Nullable.GetUnderlyingType(type);
            if (nullable != null)
                return string.IsNullOrEmpty(value) ? null : ConvertType(nullable, value);

            if (type.IsArray)
            {
                if (value == "[]" || string.IsNullOrEmpty(value)) return null;
                value = value.Trim('[', ']');
                var elementType = type.GetElementType();
                var items = value.Split(',');
                var array = Array.CreateInstance(elementType, items.Length);
                for (int i = 0; i < items.Length; i++)
                    array.SetValue(ConvertType(elementType, items[i]), i);
                return array;
            }

            if (type.GetInterface(nameof(IDictionary)) != null)
                return JsonConvert.DeserializeObject(value, type);

            if (type.GetInterface(nameof(ICollection)) != null)
            {
                var elementType = type.GetGenericArguments()[0];
                var listType = typeof(List<>).MakeGenericType(elementType);
                var list = (IList)Activator.CreateInstance(listType);
                value = value.Trim('[', ']');
                foreach (var item in value.Split(',')) list.Add(ConvertType(elementType, item));
                return list;
            }

            return value.Contains("\\n")
                ? Convert.ChangeType(value.Replace("\\n", Environment.NewLine), type)
                : Convert.ChangeType(value, type);
        }
    }

    protected virtual void OnEndParsing()
    {

    }

    public static MemberSetterDelegate CreateMemberSetter(MemberInfo member)
    {
        var targetExp = Expression.Parameter(typeof(object), "target");
        var valueExp = Expression.Parameter(typeof(object), "value");
        var targetConverted = Expression.Convert(targetExp, member.DeclaringType);

        Expression assignExp;

        if (member is FieldInfo field)
        {
            var valueConverted = Expression.Convert(valueExp, field.FieldType);
            var fieldExp = Expression.Field(targetConverted, field);
            assignExp = Expression.Assign(fieldExp, valueConverted);
        }
        else if (member is PropertyInfo prop)
        {
            var setter = prop.GetSetMethod(true);
            if (setter == null)
                throw new InvalidOperationException($"Property '{prop.Name}' has no setter.");
            var valueConverted = Expression.Convert(valueExp, prop.PropertyType);
            assignExp = Expression.Call(targetConverted, setter, valueConverted);
        }
        else
        {
            throw new NotSupportedException("Unsupported member type.");
        }

        return Expression.Lambda<MemberSetterDelegate>(assignExp, targetExp, valueExp).Compile();
    }



    public delegate void MemberSetterDelegate(object target, object value);
}