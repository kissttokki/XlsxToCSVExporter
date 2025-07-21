using Microsoft.CodeAnalysis;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace TableExporter
{
    internal partial class Program
    {
        public static Dictionary<string, List<string>> SheetInfos;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        const int COLUMN_ROW = 0;
        const int TYPE_ROW = 1;
        const int DATA_START_ROW = 5;
        const int DATA_START_COL = 1;

        static Excel.Application _ExcelApp = null;
        static Stopwatch _Watch = new Stopwatch();
        static Regex _DataRegex = new Regex(",(?=(?:(?:[^\"]*\"[^\"]*\")*[^\"]*$))");
        static uint _ExcelProcID = 0;

        public static string EXCEL_FOLDER_DIR = Environment.CurrentDirectory;


        static ConsoleEventDelegate cHandler;
        private delegate bool ConsoleEventDelegate(int eventType);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);


        static void Main(string[] args)
        {
            string sheetInfoPath = "../../table/_sheetInfo.txt";
            var t = new FileInfo(sheetInfoPath);

            if (t.Exists == false)
            {
                SheetInfos = new Dictionary<string, List<string>>();
            }
            else
            {
                string json = File.ReadAllText(t.FullName);

                try
                {
                    SheetInfos = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);
                }
                catch
                {
                    SheetInfos = new Dictionary<string, List<string>>();
                }
            }

            ///Main
            try
                {
                    Environment.CurrentDirectory = new FileInfo(System.Environment.ProcessPath).Directory.FullName;

                    CultureInfo newCulture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
                    newCulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd HH:mm:ss";
                    newCulture.DateTimeFormat.LongTimePattern = "";
                    newCulture.DateTimeFormat.PMDesignator = "";
                    newCulture.DateTimeFormat.AMDesignator = "";
                    Thread.CurrentThread.CurrentCulture = newCulture;

                    _ExcelApp = new Excel.Application();                             // 엑셀 어플리케이션 생성

                    Console.CancelKeyPress += new ConsoleCancelEventHandler(OnCancel);
                    cHandler = new ConsoleEventDelegate(ConsoleEventCallback);
                    SetConsoleCtrlHandler(cHandler, true);

                    GetWindowThreadProcessId(new IntPtr(_ExcelApp.Hwnd), out _ExcelProcID);
                    Console.WriteLine($"Created Excel Application. ProcID : {_ExcelProcID}");


                    if (string.IsNullOrEmpty(Config.Default.EXTRACT_PATH) == false)
                    {

                        DirectoryInfo exportFolder = new DirectoryInfo(Config.Default.EXTRACT_PATH);

                        if (exportFolder.Exists == false)
                        {
                            Console.WriteLine($"{exportFolder} 경로에 폴더가 없어 현재 폴더로 타겟을 지정합니다.");

                            exportFolder = new FileInfo(System.Environment.ProcessPath).Directory;
                        }
                        Config.Default.EXTRACT_TARGET = exportFolder;
                        Config.Default.EXTRACT_ENUM_TARGET = new DirectoryInfo($"{Config.Default.EXTRACT_TARGET}{Config.Default.EXTRACT_ENUM_PATH}");
                    }
                    else
                    {
                        Config.Default.EXTRACT_TARGET = new FileInfo(System.Environment.ProcessPath).Directory;
                        Config.Default.EXTRACT_ENUM_TARGET = new DirectoryInfo($"{Config.Default.EXTRACT_TARGET}{Config.Default.EXTRACT_ENUM_PATH}");
                    }

                    if (Config.Default.EXTRACT_ENUM_TARGET.Exists == false)
                        Config.Default.EXTRACT_ENUM_TARGET.Create();




                    if (args.Length == 0)
                    {
                        ExportDirectory(new FileInfo(System.Environment.ProcessPath).Directory.FullName);
                    }
                    else
                    {
                        foreach (var arg in args)
                        {
                            try
                            {
                                if (IsDirectory(arg) == true)
                                {
                                    EXCEL_FOLDER_DIR = arg;
                                    ExportDirectory(arg);
                                }
                                else
                                {

                                    EXCEL_FOLDER_DIR = new FileInfo(arg).Directory.FullName;
                                    ExportExcel(arg);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"{arg} 파싱중 에러. \r\n{ex}");
                            }
                        }
                    }
                }
                finally
                {
                    _ExcelApp?.Quit();
                    Marshal.ReleaseComObject(_ExcelApp);
                    _ExcelApp = null;

                    if (_ExcelProcID != 0)
                    {
                        System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)_ExcelProcID);
                        excelProcess.CloseMainWindow();
                        excelProcess.Refresh();
                        excelProcess.Kill();
                    }

                    Console.WriteLine("프로그램 종료.");
                    Console.ReadLine();
                }


            File.WriteAllText(sheetInfoPath, JsonConvert.SerializeObject(SheetInfos, Formatting.Indented));
        }

        private static void OnCancel(object? sender, ConsoleCancelEventArgs e)
        {
            ExitExcel();
        }

        static bool ConsoleEventCallback(int eventType)
        {
            if (eventType == 2)
            {
                ExitExcel();
            }
            return false;
        }

        static void ExitExcel()
        {
            try
            {
                if (_ExcelApp != null)
                {
                    _ExcelApp.Quit();
                    Marshal.ReleaseComObject(_ExcelApp);
                    _ExcelApp = null;
                }

                if (_ExcelProcID != 0)
                {
                    System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)_ExcelProcID);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }
            catch { }
        }


        static void ExportDirectory(string dir)
        {
            List<string> fileNames = GetExcelFiles(dir);

            Console.WriteLine($"{fileNames.Count}개 엑셀 파일 발견.");

            foreach (var arg in fileNames)
            {
                ExportExcel(arg);
            }
        }

        static void ExportExcel(string excel)
        {
            string path = excel;                              // 엑셀 파일 저장 경로

            Console.WriteLine($"{excel} 파싱");
            _ExcelApp.DisplayAlerts = false;
            var WorkBooks = _ExcelApp.Workbooks;
            var _WorkBook = WorkBooks.Open(path);                       // 워크북 열기
            Excel.Sheets WorkSheets = _WorkBook.Worksheets;

            Dictionary<int, string> columns = new Dictionary<int, string>();
            List<List<string>> datas = new List<List<string>>();
            StringBuilder csvWriter = new StringBuilder();

            _Watch.Start();

            List<string> localCsvList = new List<string>();
            List<string> serverCsvList = new List<string>();
            List<string> csvList = new List<string>();
            List<string> enumTarget = new List<string>();


            var key = new FileInfo(excel).Name;
            SheetInfos[key] = new List<string>();

            foreach (Worksheet sheet in WorkSheets)
            {
                SheetInfos[key].Add(sheet.Name);
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
                else if (sheet.Name.StartsWith("local_", StringComparison.OrdinalIgnoreCase) == true)
                {
                    localCsvList.Add(tPath);
                }
                else if (sheet.Name.StartsWith("server_", StringComparison.OrdinalIgnoreCase) == true)
                {
                    serverCsvList.Add(tPath);
                }
                else
                {
                    csvList.Add(tPath);
                }
            }

            _WorkBook?.Close(false, "", false);   // 워크북 닫기
            WorkBooks?.Close();

            if (_WorkBook != null) Marshal.ReleaseComObject(_WorkBook);
            _WorkBook = null;
            if (WorkBooks != null) Marshal.ReleaseComObject(WorkBooks);
            WorkBooks = null;


            foreach (var csvPath in localCsvList)
            {
                var sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] LocalTable Parsing CSV...");
                ParsingCSV(csvPath, "local");

                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine($"[{sheetName}] LocalTable Parsing CSV Done {_Watch.Elapsed.TotalSeconds}s");
                _Watch.Restart();
            }


            foreach (var csvPath in serverCsvList)
            {
                var sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] ServerTable Parsing CSV...");
                ParsingCSV(csvPath, "server");

                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine($"[{sheetName}] ServerTable Parsing CSV Done {_Watch.Elapsed.TotalSeconds}s");
                _Watch.Restart();
            }

            foreach (var csvPath in csvList)
            {
                var sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] Parsing CSV...");
                ParsingCSV(csvPath);

                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine($"[{sheetName}] Parsing CSV Done {_Watch.Elapsed.TotalSeconds}s");
                _Watch.Restart();
            }

            foreach (var csvPath in enumTarget)
            {
                var sheetName = Path.GetFileName(csvPath);
                Console.WriteLine($"[{sheetName}] Parsing Enum...");
                ParsingEnum(csvPath);

                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine($"[{sheetName}] Parsing Enum Done {_Watch.Elapsed.TotalSeconds}s");
                _Watch.Restart();
            }
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

        static bool IsDirectory(string path)
        {
            FileAttributes attr = File.GetAttributes(path);

            //detect whether its a directory or file
            if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                return true;
            else
                return false;
        }
    }
}
