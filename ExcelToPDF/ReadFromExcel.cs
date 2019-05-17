using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace PdfMapCreator
{
    public class ReadFromExcel
    {
        public static List<ExcelDataModel> GetExcelData(string path)
        {
            List<ExcelDataModel> excelDatas = new List<ExcelDataModel>();
            ExcelDataModel excelData = new ExcelDataModel();

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(path);
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;

            if (excelRange.Columns.Count == 3)
            {
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        {
                            if (j == 1)
                            {
                                excelData.NamePlace = excelRange.Cells[i, j].Value2.ToString();
                            }
                            try
                            {
                                if (j == 2)
                                {
                                    excelData.Latitude = Convert.ToDouble(excelRange.Cells[i, j].Value2);
                                }
                                if (j == 3)
                                {
                                    excelData.Longitude = Convert.ToDouble(excelRange.Cells[i, j].Value2);
                                }
                            }
                            catch
                            {
                                Console.WriteLine($"\n{excelData.NamePlace} does not have the appropriate coordinates!");
                            }
                        }
                    }
                    if (!excelDatas.Exists(x => x.NamePlace == excelData.NamePlace) && excelData.NamePlace != null)
                    {
                        excelDatas.Add(new ExcelDataModel { NamePlace = excelData.NamePlace, Latitude = excelData.Latitude, Longitude = excelData.Longitude });
                    }
                }

                var longestNameCount = 0;
                if (excelDatas.Count > 0)
                {
                    longestNameCount = excelDatas.Max(x => x.NamePlace.Count());
                }

                Console.WriteLine();

                foreach (var item in excelDatas)
                {
                    Console.WriteLine(item.NamePlace + string.Empty.PadLeft(longestNameCount - item.NamePlace.Count() + 2) + item.Latitude + "  " + item.Longitude);
                }

                Message.Wait();

                ////iterate over the rows and columns and print to the console as it appears in the file
                ////excel is not zero based!!
                //for (int i = 1; i <= rowCount; i++)
                //{
                //    for (int j = 1; j <= colCount; j++)
                //    {
                //        //new line
                //        if (j == 1)
                //            Console.Write("\r\n");

                //        //write the value to the console
                //        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                //            Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
                //    }
                //}
            }
            else
            {
                Message.FileFormatError();
                return null;
            }
            //cleanup
            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorksheet);

            //close and release
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorkbook);

            //quit and release
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return excelDatas;            
        }
    }
}
