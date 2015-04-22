using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using RawDataPrint;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Pdf.Printing;


namespace PrinterConsole
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("** START PRINT");
            /**************
             * 範例 1 MSDN RawPrinterHelper  (EASY WAY IF WORK)
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

            /* 範例 1 重點使用 RawPrinterHelper 發送檔案 */
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
            /*
            ProcessStartInfo psInfo = new ProcessStartInfo();
            psInfo.Arguments = String.Format(" -dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q -dNumCopies=1 -sDEVICE=ljet4 -sOutputFile=\"\\\\spool\\{0}\" \"{1}\"", "EPSON M200 Series", "adobe.pdf");
            psInfo.FileName = @"C:\Program Files (x86)\gs\gs9.16\bin\gswin32c.exe";
            psInfo.UseShellExecute = false;
            Process process = Process.Start(psInfo);
            */

            /* Using helper */
            /*
            var gh = new GhostHelper();
            gh.PrintPDF(@"C:\Program Files\gs\gs9.16\bin\gswin64c.exe", 1, "EPSON M200 Series", "adobe.pdf");
            */

            /**************
             * 範例 4 使用 PrintDocument
             * https://msdn.microsoft.com/en-us/library/system.drawing.printing.printdocument%28v=vs.100%29.aspx
             **************/
            /*
            PrintDocument document = new PrintDocument();
            document.PrintPage += document_PrintPage;
            // Configuration by Dialog
            // PrintPreviewDialog ppd = new PrintPreviewDialog();
            // ppd.Document = document;
            // ppd.ShowDialog();  

            // Configuration automatic
            document.DocumentName = "test";
            document.PrinterSettings.PrinterName = "EPSON M200 Series";
            document.PrinterSettings.PrintFileName = "adobe.pdf";
            document.PrintController = new StandardPrintController();
            document.OriginAtMargins = false;
            
            document.Print();
            */

            /**************
             * 範例 5 使用 PDFSharp (EASY WAY)
             **************/
            
            PdfFilePrinter.AdobeReaderPath = @"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe";
            PdfFilePrinter printer = new PdfFilePrinter("adobe.pdf", "EPSON M200 Series");
            printer.Print(5000); // After milliseconds close.
            


            Console.WriteLine("** DONE PRINT");
            Console.ReadKey();

            /**************
             * 其他
             **************/
            // Working with USB Device http://www.developerfusion.com/article/84338/making-usb-c-friendly/
        }

        private static void document_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString("Hello World", new Font("Arial", 10), new SolidBrush(Color.Black), new PointF(10, 10));  
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
