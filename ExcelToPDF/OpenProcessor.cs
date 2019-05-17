using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PdfMapCreator
{
    public class OpenProcessor
    {
        private static string filePath;

        public static string OpenExcelFile()
        {
            Message.OpenText();

            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Open Excel file",
                InitialDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                Filter = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xls)|*.xls"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    filePath = ofd.FileName;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Environment.Exit(0);             
            }
            Console.WriteLine(filePath);
            return filePath;
        }
    }
}
