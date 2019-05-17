using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace PdfMapCreator
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            CultureInfo.CurrentCulture = new CultureInfo("en-US", false);
            Message.Header();

            List<ExcelDataModel> excelDataModels = ReadFromExcel.GetExcelData(OpenProcessor.OpenExcelFile());

            if (excelDataModels!=null && excelDataModels.Count!=0)
            {                              
                foreach (ExcelDataModel data in excelDataModels)
                {
                    HtmlToImage.ImageCapture(data.Latitude, data.Longitude);
                }

                ThreadPool.QueueUserWorkItem(new WaitCallback(Export.ExportToPdf), excelDataModels);
                
                Message.Finished();
                Thread.Sleep(1000);
            }          
        }
    }
}
