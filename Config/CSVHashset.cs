using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableExporter
{
    public class CSVHashset
    {
        public HashSet<string> DataSheetNames { get; set; }
        public HashSet<string> EnumSheetNames { get; set; }

        public CSVHashset() { }

        public CSVHashset(List<string> dataSheets, List<string> enumSheets)
        {
            DataSheetNames = new HashSet<string>();
            EnumSheetNames = new HashSet<string>();

            string ProcDirPath = $"{new FileInfo(Environment.ProcessPath).Directory.FullName}/";

            foreach (var sheet in dataSheets)
            {
                DataSheetNames.Add(sheet.Replace(ProcDirPath,""));
            }

            foreach (var sheet in enumSheets)
            {
                EnumSheetNames.Add(sheet.Replace(ProcDirPath, ""));
            }
        }

        public CompareResult Compare(CSVHashset other)
        {
            CompareResult compareResult = new CompareResult();

            compareResult.Data = new AddedRemovedHashSet()
            {
                Added = other.DataSheetNames.Where(t => this.DataSheetNames.Contains(t) == false).ToHashSet(),
                Removed = this.DataSheetNames.Where(t => other.DataSheetNames.Contains(t) == false).ToHashSet(),
            };

            compareResult.Enum = new AddedRemovedHashSet()
            {
                Added = other.EnumSheetNames.Where(t => this.EnumSheetNames.Contains(t) == false).ToHashSet(),
                Removed = this.EnumSheetNames.Where(t => other.EnumSheetNames.Contains(t) == false).ToHashSet(),
            };


            return compareResult;
        }

        public CSVHashset Copy()
        {
            CSVHashset result = new CSVHashset();

            result.DataSheetNames = new HashSet<string>();
            result.EnumSheetNames = new HashSet<string>();

            foreach (var data in this.DataSheetNames)
            {
                result.DataSheetNames.Add(data);
            }

            foreach (var data in this.EnumSheetNames)
            {
                result.EnumSheetNames.Add(data);
            }

            return result;
        }
    }

    public class CompareResult
    {
        public AddedRemovedHashSet Data;
        public AddedRemovedHashSet Enum;
    }

    public class AddedRemovedHashSet
    {
        public HashSet<string> Added { get; set; }
        public HashSet<string> Removed { get; set; }
    }
}
