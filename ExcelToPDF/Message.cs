using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfMapCreator
{
    static class Message
    {
        public static void Header()
        {
            Console.WriteLine("==== PDF Map Creator ====");
            Console.WriteLine("\nInput an Excel file with the coordinates of your locations and create a PDF file with maps of preferred locations.");
            Console.WriteLine("An Excel file must have 3 columns: PlaceName{Text}, Latitude{decimal number}, longitude{decimal number}.\n");
        }

        public static void FileFormatError()
        {
            Console.WriteLine("\nError: The file is not supported!\n\nThe excel file must contain only 3 columns(PlaceName, Latitude, Longitude)!\nPlease correct the file data.\n");
        }

        public static void OpenText()
        {
            Console.WriteLine("Please select the Excel file!");
        }

        public static void Finished()
        {
            Console.WriteLine("\nThe task has been done.");
        }

        public static void Wait()
        {
            Console.WriteLine("\nPlease wait...");
        }
    }
}
