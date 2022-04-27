using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InteropTest
{
    class InteropConversion
    {
        public string CreatePDFusingI(string path, string exportDir)
        {
            Application app = new Application();
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            app.Visible = true;
            Document objPres = null;
            var objPresSet = app.Documents;
           

            var pdfFileName = Path.ChangeExtension(path, ".pdf");
            var pdfPath = Path.Combine(exportDir, pdfFileName);
            var timer = new Stopwatch();
            timer.Start();
            try
            {
                objPres = objPresSet.Open(path, true, true, false);
                objPres.ExportAsFixedFormat(
                    pdfPath,
                    WdExportFormat.wdExportFormatPDF,
                    false,
                    WdExportOptimizeFor.wdExportOptimizeForPrint,
                    WdExportRange.wdExportAllDocument
                );
            }
            catch
            {
                pdfPath = null;
            }
            finally
            {
                app.Visible = false;
                objPres.Close(false);
                app.Quit();
                releaseObject(objPres);
                releaseObject(app);
            }
            timer.Stop();

            TimeSpan timeTaken = timer.Elapsed;
            string timeTakenString = $@"Time taken: CreatePDFusingI :" + timeTaken.ToString(@"m\:ss\.fff");
            Console.WriteLine(timeTakenString);
            return pdfPath;
        }

        public void convertDOCtoPDF(string path, string exportDir)
        {

            object misValue = System.Reflection.Missing.Value;
            var pdfPath = Path.Combine(exportDir, $"{Guid.NewGuid()}.pdf");
            string PATH_APP_PDF = pdfPath;
            Console.WriteLine($"Converting Document Path : {path}");
            var WORD = new Application();
            Document doc = null;
            var timer = new Stopwatch();
            timer.Start();
            try
            {

                doc = WORD.Documents.Open(path, ReadOnly: true);
                doc.Activate();
                doc.SaveAs2(@PATH_APP_PDF, WdSaveFormat.wdFormatPDF, misValue, misValue, misValue,
                misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            }
            finally
            {
                doc.Close(false);
                WORD.Quit();
                releaseObject(doc);
                releaseObject(WORD);
            }
          
            timer.Stop();

            TimeSpan timeTaken = timer.Elapsed;
            string timeTakenString = $@"Time taken: convertDOCtoPDF :" + timeTaken.ToString(@"m\:ss\.fff");
            Console.WriteLine(timeTakenString);

        }

        public void convertDOCtoPDFusingStream(string path, string exportDir)
        {

            object misValue = System.Reflection.Missing.Value;
            var pdfPath = Path.Combine(exportDir, $"{Guid.NewGuid()}.pdf");
            string PATH_APP_PDF = pdfPath;

            var WORD = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;
            var timer = new Stopwatch();
            timer.Start();
            try
            {

                doc = WORD.Documents.Open(path, ReadOnly: true);
                doc.Activate();
                doc.SaveAs2(@PATH_APP_PDF, WdSaveFormat.wdFormatPDF, misValue, misValue, misValue,
                misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            }
            finally
            {
                doc.Close(false);
                WORD.Quit();
                releaseObject(doc);
                releaseObject(WORD);
            }

            timer.Stop();

            TimeSpan timeTaken = timer.Elapsed;
            string timeTakenString = $@"Time taken: " + timeTaken.ToString(@"m\:ss\.fff");
            Console.WriteLine(timeTakenString);

        }
        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                //TODO
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
