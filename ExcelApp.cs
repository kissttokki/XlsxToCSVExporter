using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using static TableExporter.CSVReader;

namespace TableExporter
{
    internal class ExcelApp : IDisposable
    {
        private const string PATH_OUTPUT_CLIENT = "output/client";
        private const string PATH_OUTPUT_SERVER = "output/server";

        private const string PATH_OUTPUT_CLIENT_CSV = $"{PATH_OUTPUT_CLIENT}/csv";
        private const string PATH_OUTPUT_SERVER_CSV = $"{PATH_OUTPUT_SERVER}/csv";

        private const string PATH_OUTPUT_CLIENT_SCRIPT = $"{PATH_OUTPUT_CLIENT}/script";
        private const string PATH_OUTPUT_SERVER_SCRIPT = $"{PATH_OUTPUT_SERVER}/script";

        private const string PATH_OUTPUT_CLIENT_SCRIPT_ENUM = $"{PATH_OUTPUT_CLIENT_SCRIPT}/enum";
        private const string PATH_OUTPUT_SERVER_SCRIPT_ENUM = $"{PATH_OUTPUT_SERVER_SCRIPT}/enum";



        private Stopwatch _stopWatch = new Stopwatch();
        private Application? _excelApp;
        private uint _excelProcID = 0;


        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);


        public ExcelApp()
        {
            _excelApp = new Microsoft.Office.Interop.Excel.Application();

            GetWindowThreadProcessId(new IntPtr(_excelApp.Hwnd), out _excelProcID);
            Console.WriteLine($"Created Excel Application. ProcID : {_excelProcID}");
        }



        public void ImportDirectory(string dir)
        {
            List<string> fileNames = GetExcelFiles(dir);

            Console.WriteLine($"{fileNames.Count}개 엑셀 파일 발견.");

            foreach (var arg in fileNames)
            {
                ImportExcel(arg);
            }
        }

        public void ImportExcel(string filePath)
        {
            if (_excelApp == null)
            {
                Console.WriteLine($"[ERR] ExcelApp is not initialized");
                return;
            }


            Console.WriteLine($"{filePath} 파싱");
            _excelApp.DisplayAlerts = false;
            var WorkBooks = _excelApp.Workbooks;
            var _WorkBook = WorkBooks.Open(filePath);                       // 워크북 열기
            Sheets WorkSheets = _WorkBook.Worksheets;

            Dictionary<int, string> columns = new Dictionary<int, string>();
            List<List<string>> datas = new List<List<string>>();
            StringBuilder csvWriter = new StringBuilder();

            _stopWatch.Start();

            List<string> tableList = new List<string>();
            List<string> enumList = new List<string>();
            string ProcDirPath = new FileInfo(Environment.ProcessPath).Directory.FullName;


            var key = new FileInfo(filePath).Name;

            Directory.CreateDirectory($"{ProcDirPath}/output");

            foreach (Worksheet sheet in WorkSheets)
            {
                if (sheet.Name.StartsWith("_") == true)
                {
                    continue;
                }

                var tPath = $"{ProcDirPath}/output/{sheet.Name}.csv";

                sheet.SaveAs(tPath, XlFileFormat.xlCSVUTF8, Local: true);

                if (sheet.Name.StartsWith("enum_", StringComparison.OrdinalIgnoreCase) == true)
                {
                    enumList.Add(tPath);
                }
                else
                {
                    tableList.Add(tPath);
                }
            }

            _WorkBook?.Close(false, "", false);   // 워크북 닫기
            WorkBooks?.Close();

            if (_WorkBook != null) Marshal.ReleaseComObject(_WorkBook);
            _WorkBook = null;
            if (WorkBooks != null) Marshal.ReleaseComObject(WorkBooks);
            WorkBooks = null;


            if (tableList.Count > 0)
            {
                Console.WriteLine($"[BaseTable] Generate BaseTable CS...");

                var codes = File.ReadAllText("./BaseTableTemplate.cs");

                FileExtension.ProcessCreateFile(PATH_OUTPUT_CLIENT_SCRIPT, "BaseTable.cs", codes);
                FileExtension.ProcessCreateFile(PATH_OUTPUT_SERVER_SCRIPT, "BaseTable.cs", codes);
            }



            if (Config.Default.SaveHistories == null)
                Config.Default.SaveHistories = new Dictionary<string, CSVHashset>();

            var fileName = Path.GetFileName(filePath);


            var tempHistory = new CSVHashset(tableList, enumList);


            foreach (var csvPath in tableList)
            {
                CSVReader clientCSV = new CSVReader(csvPath, CSVWriteType.Client);
                CSVReader serverCSV = new CSVReader(csvPath, CSVWriteType.Server);

                string sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] Parsing CSV...");

                tempHistory.DataSheetNames.Add(FileExtension.ProcessCreateFile(PATH_OUTPUT_CLIENT_CSV, sheetName, clientCSV.GetCSV()));
                tempHistory.DataSheetNames.Add(FileExtension.ProcessCreateFile(PATH_OUTPUT_SERVER_CSV, sheetName, serverCSV.GetCSV()));

                string csName = sheetName.Replace(".csv", ".cs");

                tempHistory.DataSheetNames.Add(FileExtension.ProcessCreateFile(PATH_OUTPUT_CLIENT_SCRIPT, csName, clientCSV.GetClassCode()));
                tempHistory.DataSheetNames.Add(FileExtension.ProcessCreateFile(PATH_OUTPUT_SERVER_SCRIPT, csName, serverCSV.GetClassCode()));

                Console.WriteLine($"[{sheetName}] Parsing CSV Done {_stopWatch.Elapsed.TotalSeconds}s");
                _stopWatch.Restart();
            }

            foreach (var csvPath in enumList)
            {
                CSVReader clientCSV = new CSVReader(csvPath, CSVWriteType.Client);
                CSVReader serverCSV = new CSVReader(csvPath, CSVWriteType.Server);

                string sheetName = Path.GetFileName(csvPath);

                Console.WriteLine($"[{sheetName}] Parsing Enum...");
                foreach (var enumCode in clientCSV.GetEnumCodes())
                {
                    tempHistory.EnumSheetNames.Add(FileExtension.ProcessCreateFile(PATH_OUTPUT_CLIENT_SCRIPT_ENUM, $"{enumCode.name}.cs", enumCode.code));
                }

                foreach (var enumCode in serverCSV.GetEnumCodes())
                {
                    tempHistory.EnumSheetNames.Add(FileExtension.ProcessCreateFile(PATH_OUTPUT_SERVER_SCRIPT_ENUM, $"{enumCode.name}.cs", enumCode.code));
                }

                Console.WriteLine($"[{sheetName}] Parsing Enum Done {_stopWatch.Elapsed.TotalSeconds}s");
                _stopWatch.Restart();
            }


            if(string.IsNullOrEmpty(Config.Default.XlsxTargetsFolder) == false)
                File.WriteAllText($"{Config.Default.XlsxTargetsFolder}/_sheetInfos.txt", JsonConvert.SerializeObject(Config.Default.SaveHistories, Formatting.Indented));



            if (Config.Default.SaveHistories.TryGetValue(fileName, out var hashSet) == false || hashSet == null)
            {
                hashSet = tempHistory;
            }
            else
            {
                var compare = hashSet.Compare(tempHistory);
                foreach (var item in compare.Data.Removed)
                {
                    Console.WriteLine($"!!!! Delete - {item}");
                    File.Delete(item);
                }
                foreach (var item in compare.Enum.Removed)
                {
                    Console.WriteLine($"!!!! Delete - {item}");
                    File.Delete(item);
                }
            }

            Config.Default.SaveHistories[fileName] = tempHistory;

            Config.Default.Save();
        }

        public void MoveFiles()
        {
            Console.WriteLine("-----MOVE FILES----");
            string procDir = new FileInfo(Environment.ProcessPath).Directory.FullName;

            Console.WriteLine($"{procDir}/{PATH_OUTPUT_CLIENT_CSV}");
            CopyFolder($"{procDir}/{PATH_OUTPUT_CLIENT_CSV}", Config.Default.OutputClientCSVDir);
            CopyFolder($"{procDir}/{PATH_OUTPUT_SERVER_CSV}", Config.Default.OutputServerCSVDir);
            CopyFolder($"{procDir}/{PATH_OUTPUT_CLIENT_SCRIPT}", Config.Default.OutputClientCsharpScriptDir);
            CopyFolder($"{procDir}/{PATH_OUTPUT_SERVER_SCRIPT}", Config.Default.OutputServerCsharpScriptDir);
            
            Console.WriteLine("-----COMPLETE----");
        }


        private void CopyFolder(string sourceDir, string targetDir)
        {
            if (string.IsNullOrEmpty(targetDir)) return;

            if (Directory.Exists(targetDir))
            {
                foreach (var file in Directory.GetFiles(targetDir, "*.*", SearchOption.AllDirectories))
                {
                    // .meta 파일은 건너뛰기
                    if (Path.GetExtension(file).Equals(".meta", StringComparison.OrdinalIgnoreCase))
                        continue;

                    File.Delete(file);
                }
            }
            else
            {
                Directory.CreateDirectory(targetDir);
            }



            foreach (string dirPath in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceDir, targetDir));
            }

            foreach (string filePath in Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories))
            {
                string destFile = filePath.Replace(sourceDir, targetDir);
                File.Copy(filePath, destFile, overwrite: true);

                Console.WriteLine($"-- Copy {filePath} -> {destFile}");
            }



            foreach (var dir in Directory.GetDirectories(targetDir, "*", SearchOption.AllDirectories))
            {
                // 폴더 내부에 .meta만 남은 경우 삭제
                if (!Directory.EnumerateFileSystemEntries(dir).Any(f => !f.EndsWith(".meta", StringComparison.OrdinalIgnoreCase)))
                    Directory.Delete(dir, recursive: true);
            }
        }




        //private bool ChecksumCSV(string filePath)
        //{
        //    using var stream = File.OpenRead(filePath);
        //    using var sha = SHA256.Create();
        //    byte[] hash = sha.ComputeHash(stream);
        //    string checksum = BitConverter.ToString(hash).Replace("-", "");


        //    return string.Equals(savedChecksum, checksum) == true;
        //}
        static List<string> GetExcelFiles(string directoryPath)
        {
            List<string> result = new List<string>();
            var dirInfo = new DirectoryInfo(directoryPath);
            var files = dirInfo.GetFiles();

            foreach (var f in files)
            {
                if (f.Extension.StartsWith(".xl") == true)
                {
                    if (f.Name.StartsWith("~") == false) result.Add(f.FullName);
                }
            }

            return result;
        }



        public void Dispose()
        {
            ExitExcel();
        }


        private void ExitExcel()
        {
            try
            {
                if (_excelApp != null)
                {
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                    _excelApp = null;
                }

                if (_excelProcID != 0)
                {
                    System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)_excelProcID);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }
            catch { }
        }
    }
}
