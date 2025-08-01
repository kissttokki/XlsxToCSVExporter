using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace TableExporter
{
    internal class ExcelApp : IDisposable
    {
        private Stopwatch _stopWatch = new Stopwatch();
        private Application? _excelApp;
        private uint _excelProcID = 0;
        private Dictionary<string, List<string>> _sheetInfos = new Dictionary<string, List<string>>();



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
            List<string> enumTarget = new List<string>();


            var key = new FileInfo(filePath).Name;
            _sheetInfos[key] = new List<string>();

            foreach (Worksheet sheet in WorkSheets)
            {
                _sheetInfos[key].Add(sheet.Name);
                if (sheet.Name.StartsWith("_") == true)
                {
                    continue;
                }

                var tPath = $"{Path.GetTempPath()}{sheet.Name}.csv";

                sheet.SaveAs(tPath, XlFileFormat.xlCSVUTF8, Local: true);

                if (sheet.Name.StartsWith("enum_", StringComparison.OrdinalIgnoreCase) == true)
                {
                    enumTarget.Add(tPath);
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

                FileExtension.ProcessCreateFile(Config.Default.OutputClientCsharpScriptDir, "BaseTable.cs", codes, "/client");
                FileExtension.ProcessCreateFile(Config.Default.OutputServerCsharpScriptDir, "BaseTable.cs", codes, "/server");
            }


            foreach (var csvPath in tableList)
            {
                var sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] Parsing CSV...");
                Processing.ParsingCSV(csvPath);

                Console.WriteLine($"[{sheetName}] Parsing CSV Done {_stopWatch.Elapsed.TotalSeconds}s");
                _stopWatch.Restart();
            }

            foreach (var csvPath in enumTarget)
            {
                var sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] Parsing Enum...");
                Processing.GenerateEnum(csvPath);
                Console.WriteLine($"[{sheetName}] Parsing Enum Done {_stopWatch.Elapsed.TotalSeconds}s");
                _stopWatch.Restart();
            }




            File.WriteAllText($"{Config.Default.XlsxTargetsFolder}/_sheetInfos.txt", JsonConvert.SerializeObject(_sheetInfos, Formatting.Indented));
        }


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
