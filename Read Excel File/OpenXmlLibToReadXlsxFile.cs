using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Data;
using System.IO;


namespace Read_Excel_File
{
    public class OpenXmlLibToReadXlsxFile
    {



        ///  Read Data from selected excel file into DataTable 
        /// </summary> 
        /// <param name="filename">Excel File Path</param> 
        /// <returns></returns> 
        /// 

        
        public DataTable ReadExcelFile(string filename)
        {
            // Initialize an instance of DataTable 
            DataTable dt = new DataTable();


            try
            {
                // Use SpreadSheetDocument class of Open XML SDK to open excel file 
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, true))
                {
                    // Get Workbook Part of Spread Sheet Document 
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;


                    // Get all sheets in spread sheet document  
                    IEnumerable<Sheet> sheetcollection = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();


                    // Get relationship Id 
                    //  string relationshipId = sheetcollection.First().Id.Value; 

                    var relationshipIdCollection = sheetcollection.ToList();

                    foreach (var item in relationshipIdCollection)
                    {
                        string relationshipId = item.Id.Value;


                        // Get sheet1 Part of Spread Sheet Document 
                        WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);


                        // Get Data in Excel file 
                        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                        IEnumerable<Row> rowcollection = sheetData.Descendants<Row>();


                        if (rowcollection.Count() == 0)
                        {
                            continue;
                        }


                        // Add columns 
                        foreach (Cell cell in rowcollection.ElementAt(0))
                        {
                            dt.Columns.Add(GetValueOfCell(spreadsheetDocument, cell));
                        }
                        try
                        {
                            //if (string.Compare(dt.Columns[0].ColumnName, "Name", true) == 0 && string.Compare(dt.Columns[1].ColumnName.Trim(), "Class", true) == 0 && string.Compare(dt.Columns[2].ColumnName, "Roll Number", true) == 0)
                            //{
                                // Add rows into DataTable 
                                foreach (Row row in rowcollection)
                                {
                                    DataRow temprow = dt.NewRow();
                                    int columnIndex = 0;
                                    foreach (Cell cell in row.Descendants<Cell>())
                                    {
                                        // Get Cell Column Index 
                                        int cellColumnIndex = GetColumnIndex(GetColumnName(cell.CellReference));


                                        if (columnIndex < cellColumnIndex)
                                        {
                                            do
                                            {
                                                temprow[columnIndex] = string.Empty;
                                                columnIndex++;
                                            }


                                            while (columnIndex < cellColumnIndex);
                                        }


                                        temprow[columnIndex] = GetValueOfCell(spreadsheetDocument, cell);
                                        columnIndex++;
                                    }


                                    // Add the row to DataTable 
                                    // the rows include header row 
                                    dt.Rows.Add(temprow);
                                }
                                dt.Rows.RemoveAt(0);
                                break;
                            //}




                            //else
                            //{
                            //    dt.Columns.Clear();
                            //}


                        }

                        catch (Exception e)
                        {
                            dt = null;


                        }

                    }

                }
                // Here remove header row 

                return dt;

            }
            catch (IOException ex)
            {
                throw new IOException(ex.Message);
            }
        }
        private int GetColumnIndex(string columnName)
        {
            int columnIndex = 0;
            int factor = 1;

            // From right to left
            for (int position = columnName.Length - 1; position >= 0; position--)
            {
                // For letters
                if (Char.IsLetter(columnName[position]))
                {
                    columnIndex += factor * ((columnName[position] - 'A') + 1) - 1;
                    factor *= 26;
                }
            }

            return columnIndex;
        }
        private string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name of cell
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }


        /// <summary> 
        ///  Get Value of Cell  
        /// </summary> 
        /// <param name="spreadsheetdocument">SpreadSheet Document Object</param> 
        /// <param name="cell">Cell Object</param> 
        /// <returns>The Value in Cell</returns> 
        /// 

         
        private static string GetValueOfCell(SpreadsheetDocument spreadsheetdocument, Cell cell)
        {
            // Get value in Cell 
            SharedStringTablePart sharedString = spreadsheetdocument.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return string.Empty;
            }


            string cellValue = cell.CellValue.InnerText;

            // The condition that the Cell DataType is SharedString 
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedString.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
            }
            else
            {
                return cellValue;
            }
        } 



    }
}