using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace Read_Excel_File
{
    class Program
    {
        static void Main(string[] args)
        {
            var xmlObj = new OpenXmlLibToReadXlsxFile();

/// Required path of Excel file
            DataTable table = xmlObj.ReadExcelFile(@"C:\Users\Administrator\Downloads\Sample-2.xlsx");
           // ConvertDocToHTML.ConvertTOhtml();
            ConvertToHTMLWithImages.ConVertDOcWithImages();
        }
    }
}
