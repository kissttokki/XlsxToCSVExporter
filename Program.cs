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
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);

        private static ExcelApp? s_excelApp;
        private static ConsoleEventDelegate? s_cHandler;
        private delegate bool ConsoleEventDelegate(int eventType);


        static void Main(string[] args)
        {
            CultureInfo newCulture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            newCulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd HH:mm:ss";
            newCulture.DateTimeFormat.LongTimePattern = "";
            newCulture.DateTimeFormat.PMDesignator = "";
            newCulture.DateTimeFormat.AMDesignator = "";
            Thread.CurrentThread.CurrentCulture = newCulture;
            ///Main
            try
            {
                Console.CancelKeyPress += new ConsoleCancelEventHandler(OnCancel);
                s_cHandler = new ConsoleEventDelegate(ConsoleEventCallback);
                SetConsoleCtrlHandler(s_cHandler, true);

                s_excelApp = new ExcelApp();




                if (args.Length == 0)
                {
                    s_excelApp.ImportDirectory(new FileInfo(System.Environment.ProcessPath).Directory.FullName);
                }
                else
                {
                    foreach (var arg in args)
                    {
                        try
                        {
                            if (IsDirectory(arg) == true)
                            {
                                s_excelApp.ImportDirectory(arg);
                                //EXCEL_FOLDER_DIR = arg;
                                //ExportDirectory(arg);
                            }
                            else
                            {
                                s_excelApp.ImportExcel(arg);
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
                s_excelApp?.Dispose();
                s_excelApp = null;
                Console.WriteLine("프로그램 종료.");
                Console.ReadLine();
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

        private static void OnCancel(object? sender, ConsoleCancelEventArgs e)
        {
            s_excelApp?.Dispose();
            s_excelApp = null;
        }

        static bool ConsoleEventCallback(int eventType)
        {
            if (eventType == 2)
            {
                s_excelApp?.Dispose();
                s_excelApp = null;
            }
            return false;
        }
    }
}
