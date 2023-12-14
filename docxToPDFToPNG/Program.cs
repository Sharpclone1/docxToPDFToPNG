using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using ImageMagick;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace ConsoleApplication1
{
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            var openFileDialog = new OpenFileDialog { Filter = "Word Documents|*.docx" };
            var wordFilename = openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : string.Empty;
            Console.WriteLine(wordFilename);
            var pdfFileDir = Path.GetDirectoryName(wordFilename);
            Console.WriteLine(pdfFileDir);
            var pdfFilename = Path.Combine(pdfFileDir, "welcome.pdf");
            Console.WriteLine(pdfFilename);

            var wordApp = new Application { Visible = false };
            var doc = wordApp.Documents.Open(wordFilename);
            doc.SaveAs2(pdfFilename, WdSaveFormat.wdFormatPDF);
            doc.Close();
            wordApp.Quit();
            Thread.Sleep(2000);
            ConvertPdfToPngWithMagick(pdfFilename, pdfFileDir);
            File.Delete(pdfFilename);
        }

        private static void ConvertPdfToPngWithMagick(string pdfPath, string outputPath)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            baseDir = Path.GetFullPath(baseDir); // Resolves any ".." in the path

            MagickNET.SetGhostscriptDirectory(baseDir);

            using (var images = new MagickImageCollection())
            {
                images.Read(pdfPath);
                images.Write(Path.Combine(outputPath, "Welcome.png"));
            }
        }
    }
}