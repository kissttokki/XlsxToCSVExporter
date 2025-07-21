using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace TableExporter
{
    internal class Config
    {
        public string EXTRACT_CLASS_PATH { get; set; }

        public string EXTRACT_PATH { get; set; }

        public string EXTRACT_ENUM_PATH { get; set; } = "/enum";

        [JsonIgnore]
        public DirectoryInfo EXTRACT_TARGET { get; set; }

        [JsonIgnore]
        public DirectoryInfo EXTRACT_ENUM_TARGET { get; set; }



        public static Config Default
        {
            get
            {
                if (_Default == null)
                {
                    var info = new FileInfo($"{new FileInfo(System.Environment.ProcessPath).Directory.FullName}/config.json");

                    if (info.Exists == true)
                    {
                        _Default = JsonConvert.DeserializeObject<Config>(File.ReadAllText($"{new FileInfo(System.Environment.ProcessPath).Directory.FullName}/config.json"));
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
    }
}
