using DocumentsMerger.Models;
using System;
using System.Threading;
using System.Collections.Generic;
using System.Web.Http;
using iText;
using System.Net.Http;
using System.Net;
using System.IO;
using System.Web.Mvc;
using System.Net.Http.Headers;
using System.Web;
using System.Runtime.InteropServices;

namespace DocumentsMerger.Controllers
{
    public class MergeController : ApiController
    {

        public iText.Kernel.Pdf.PdfDocument resultPdf;

        public string returnResult = "";

        public const string resultPath = "C:\\ResultPrints\\";

        public string remoteHost;


        // POST api/merge 
        [System.Web.Http.Route("api/merge"), System.Web.Http.HttpPost]
        public IHttpActionResult MergeDocs(DocumentsData documentsData)
        {
            

            System.IO.Directory.CreateDirectory(resultPath + documentsData.unique_filepath);

            this.resultPdf = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfWriter(resultPath + documentsData.unique_filepath + "\\resultMerge" + documentsData.unique_filepath + ".pdf"));
            this.returnResult = resultPath + documentsData.unique_filepath + "\\resultMerge" + documentsData.unique_filepath + ".pdf";

            this.remoteHost = HttpContext.Current.Request.Url.Host.ToString().Trim();

            foreach (string filename in documentsData.filenames)
            {

                string[] filename_parts = filename.Split('.');

                string sourceFilePath = this.remoteHost + "/" + documentsData.filepath + "/" + filename;

                if (filename_parts[1] == "docx")
                {

                    Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

                    var docs = app.Documents;
                    var doc = docs.Open(sourceFilePath, true, true, false); 

                    doc.Activate();

                    string printPath = resultPath + documentsData.unique_filepath + "\\" + filename_parts[0] + ".pdf";

                    doc.ExportAsFixedFormat(printPath.ToString(), Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                    doc.Close();

                    Marshal.FinalReleaseComObject(doc);

                    Marshal.FinalReleaseComObject(docs);

                    app.Quit();

                    Marshal.FinalReleaseComObject(app);

                    app = null; doc = null; docs = null;

                    GC.Collect(); GC.WaitForPendingFinalizers();
                    GC.Collect(); GC.WaitForPendingFinalizers();

                    this.GenerateSinglePdf(printPath);

                }
                else if (filename_parts[1] == "xlsx")
                {

                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                    var wbs = app.Workbooks;
                    var wb = wbs.Open(sourceFilePath);

                    wb.Activate();

                    string printPath = resultPath + documentsData.unique_filepath + "\\" + filename_parts[0] + ".pdf";

                    wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, printPath.ToString());

                    wb.Close();
                    
                    Marshal.FinalReleaseComObject(wb);

                    wbs.Close();

                    Marshal.FinalReleaseComObject(wbs);

                    app.Quit();

                    Marshal.FinalReleaseComObject(app);

                    app = null; wb = null; wbs = null;

                    GC.Collect(); GC.WaitForPendingFinalizers();
                    GC.Collect(); GC.WaitForPendingFinalizers();

                    this.GenerateSinglePdf(printPath);

                }
                else
                {
                    throw new Exception("Wrong file format. Only docx/xlsx 2007-2013/16 file formats are supported.");
                }

            }

            this.resultPdf.Close();

            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();

            byte[] byteArray = File.ReadAllBytes(this.returnResult);
/*            MemoryStream ms = new MemoryStream();
            ms.Write(byteArray, 0 , byteArray.Length);                 // FOR LARGER FILES SEND FILESTREAM CONTENT TO THE RESPONSE IF NEEDED
            ms.Position = 0;*/

            IHttpActionResult response;

            HttpResponseMessage responseMsg = new HttpResponseMessage(HttpStatusCode.OK);


            responseMsg.Content = new ByteArrayContent(byteArray);

            responseMsg.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");

            responseMsg.Content.Headers.ContentDisposition.FileName = "resultMerge_" + documentsData.unique_filepath + ".pdf";

            responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            response = ResponseMessage(responseMsg);


            return response;

        }


        private void GenerateSinglePdf(string targetPDF)
        {

            iText.Kernel.Pdf.PdfDocument docToMerge = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(targetPDF));

            docToMerge.CopyPagesTo(1, docToMerge.GetNumberOfPages(), this.resultPdf);

        }

    }
}
