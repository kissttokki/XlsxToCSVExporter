using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace TableExporter
{
    internal class Config
    {
        public static Config Default
        {
            get
            {
                if (_Default == null)
                {
                    var info = new FileInfo($"{new FileInfo(Environment.ProcessPath).Directory.FullName}/config.json");

                    if (info.Exists == true)
                    {
                        _Default = JsonConvert.DeserializeObject<Config>(File.ReadAllText($"{new FileInfo(Environment.ProcessPath).Directory.FullName}/config.json"));
                    }
                    else
                    {
                        _Default = new Config();
                        File.WriteAllText(info.FullName, JsonConvert.SerializeObject(_Default, Formatting.Indented));
                    }

                }

                return _Default;
            }
        }
        private static Config _Default;



        [JsonIgnore]
        public Regex DataRegex { get; private set; } = new Regex(",(?=(?:(?:[^\"]*\"[^\"]*\")*[^\"]*$))");

        public int ColumnNameRow { get; set; } = 0;
        public int DataTypeRow { get; set; } = 1;
        public int BuildTargetRow { get; set; } = 2;
        public int DescriptionRow { get; set; } = 3;
        public int DataRow { get; set; } = 4;

        public int EnumColumnNameRow { get; set; } = 0;
        public int EnumBuildTargetRow { get; set; } = 1;
        public int EnumDataRow { get; set; } = 3;


        public string XlsxTargetsFolder { get; set; }
        
        public string OutputServerCSVDir { get; set; }
        public string OutputClientCSVDir { get; set; }
        public string OutputServerCsharpScriptDir { get; set; }
        public string OutputClientCsharpScriptDir { get; set; }
        public string ResultNameSpace { get; set; }

        public Dictionary<string, CSVHashset> SaveHistories { get; set; }

        public void Save()
        {
            var path = $"{new FileInfo(Environment.ProcessPath).Directory.FullName}/config.json";
            File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
        }

    }
}
