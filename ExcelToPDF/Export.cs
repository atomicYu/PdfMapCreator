using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;

namespace PdfMapCreator
{
    public class Export
    {
        public static void ExportToPdf(object exD)
        {
            List<ExcelDataModel> excelDatas = exD as List<ExcelDataModel>;
            try
            {
                Document pdfDoc = new Document(PageSize.A4, 30f, 30f, 20f, 20f);
                var writer = PdfWriter.GetInstance(pdfDoc, new FileStream($"{Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)}/Maps.pdf", FileMode.Create));

                pdfDoc.Open();

                for (int i = 0; i < excelDatas.Count; i++)
                {
                    PdfContentByte cb = writer.DirectContent;
                    Rectangle rect = new Rectangle(10f, 20f, 585f, 822f)
                    {
                        Border = Rectangle.BOX,
                        BorderWidth = 0.7f,
                        BorderColor = new BaseColor(0, 0, 0)
                    };
                    cb.Rectangle(rect);

                    //cb.Stroke();

                    var spacer = new Paragraph("")
                    {
                        SpacingBefore = 10f,
                        SpacingAfter = 10f
                    };

                    pdfDoc.Add(spacer);
                    pdfDoc.Add(spacer);

                    var pFont1 = FontFactory.GetFont(Font.FontFamily.TIMES_ROMAN.ToString(), 20, Font.NORMAL, BaseColor.BLACK);

                    Paragraph paragraph1 = new Paragraph("Location Map", pFont1)
                    {
                        Alignment = Element.ALIGN_CENTER
                    };
                    pdfDoc.Add(paragraph1);
                    pdfDoc.Add(spacer);
                    pdfDoc.Add(spacer);

                    var tableFont = FontFactory.GetFont(Font.FontFamily.TIMES_ROMAN.ToString(), 12, Font.NORMAL, BaseColor.BLACK);

                    PdfPTable tableHead = new PdfPTable(2);
                    tableHead.DefaultCell.Padding = 3;
                    float[] headerWidths = { 0.6f, 0.8f };
                    tableHead.SetWidths(headerWidths);
                    tableHead.WidthPercentage = 50;
                    tableHead.HorizontalAlignment = 0;

                    tableHead.DefaultCell.BorderColor = BaseColor.BLACK;
                    tableHead.DefaultCell.BorderWidth = 0.7f;
                    tableHead.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
                    tableHead.AddCell(new Phrase("Location Name:", tableFont));
                    tableHead.AddCell(excelDatas[i].NamePlace);
                    tableHead.AddCell("Latitude:");
                    tableHead.AddCell(excelDatas[i].Latitude.ToString());
                    tableHead.AddCell("Longitude:");
                    tableHead.AddCell(excelDatas[i].Longitude.ToString());
                    pdfDoc.Add(tableHead);

                    pdfDoc.Add(spacer);

                    PdfPTable tableMap = new PdfPTable(1);
                    tableMap.DefaultCell.Padding = 0.6f;
                    tableMap.WidthPercentage = 100;
                    tableMap.HorizontalAlignment = 0;
                    tableMap.DefaultCell.BorderColor = BaseColor.BLACK;
                    tableMap.DefaultCell.BorderWidth = 0.7f;
                    tableMap.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;

                    
                    //HtmlToImage.DocumentCompleted();
                    System.Drawing.Image image = HtmlToImage.Img[i];
                    var iTextImage1 = Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Jpeg);
                    tableMap.AddCell(iTextImage1);
                    pdfDoc.Add(tableMap);

                    pdfDoc.NewPage();
                }
                pdfDoc.Close();
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
