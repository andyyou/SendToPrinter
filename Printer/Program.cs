using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using RawDataPrint;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;


namespace PrinterConsole
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("** START PRINT");
            /**************
             * 範例 1 MSDN RawPrinterHelper (Failure)
             **************/
            foreach (var printername in PrinterSettings.InstalledPrinters)
            {
                Console.WriteLine(printername); // 列出所有的印表機
            }

            /* Options
            // 讓使用者挑機器(要用這個需要用 System.Form)
            PrintDialog pd = new PrintDialog();
            pd.Document = new PrintDocument ();
            pd.PrinterSettings = new PrinterSettings ();
            pd.UseEXDialog = true;
            if (DialogResult.OK == pd.ShowDialog (this)) {
                // 傳送文字到印表機
                RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, s);
            }
            */

            /* 發送檔案 */
            // RawPrinterHelper.SendFileToPrinter(new PrinterSettings().PrinterName, "adobe.pdf");

            /**************
             * 範例 2 使用 Adobe Reader (Failure)
             * 需要安裝 Acrobat Reader
             **************/
            /*
            ProcessStartInfo info = new ProcessStartInfo();
            info.Verb = "print";
            info.FileName = @"c:\adobe.pdf";
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;

            Process process = new Process();
            process.StartInfo = info;
            process.Start();

            process.WaitForInputIdle();
            System.Threading.Thread.Sleep(3000);
            if (false == process.CloseMainWindow())
                process.Kill();
           */

            /**************
             * 範例 3 使用 Ghostscript (Failure)
             * 需安裝 Ghostscript http://www.ghostscript.com/doc/9.16/Readme.htm
             **************/
            // var gh = new GhostHelper();
            // gh.PrintPDF(@"C:\Program Files\gs\gs9.16\bin\gswin64c.exe", 1, "EPSON M200 Series", "adobe.pdf");

            Console.WriteLine("** DONE PRINT");
            Console.ReadKey();
        }

        class GhostHelper
        {
            /// <summary>
            /// Prints the PDF.
            /// </summary>
            /// <param name="ghostScriptPath">The ghost script path. Eg "C:\Program Files\gs\gs8.71\bin\gswin32c.exe"</param>
            /// <param name="numberOfCopies">The number of copies.</param>
            /// <param name="printerName">Name of the printer. Eg \\server_name\printer_name</param>
            /// <param name="pdfFileName">Name of the PDF file.</param>
            /// <returns></returns>
            public bool PrintPDF(string ghostScriptPath, int numberOfCopies, string printerName, string pdfFileName)
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.Arguments = " -dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q -dNumCopies=" + Convert.ToString(numberOfCopies) + " -sDEVICE=ljet4 -sOutputFile=\"\\\\spool\\" + printerName + "\" \"" + pdfFileName + "\" ";
                startInfo.FileName = ghostScriptPath;
                startInfo.UseShellExecute = false;

                startInfo.RedirectStandardError = true;
                startInfo.RedirectStandardOutput = true;

                Process process = Process.Start(startInfo);

                Console.WriteLine(process.StandardError.ReadToEnd() + process.StandardOutput.ReadToEnd());

                process.WaitForExit(30000);
                if (process.HasExited == false) process.Kill();


                return process.ExitCode == 0;
            }
        }
    }
}
