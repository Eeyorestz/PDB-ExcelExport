using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML;
using ClosedXML.Excel;
using ExcelDataReader;




namespace PDB_Excel_Data_Extractor
{
    public class ExcelReader
    {

        /// <summary>
        /// Reads the data from Excel table and transform it into DataTable.
        /// </summary>
        /// <param name="fileName">The name of the Excel file</param>
        /// <param name="sheetName">The name of the worksheet in the Excel file</param>
        /// <returns></returns>
        /// 
        /// 
        /// 


        public  DataTable ExcelToDataTable(string sheetName, string filePath)
        {
            DataSet result = new DataSet();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    reader.Read();
                    result = reader.AsDataSet();
                }
            }
            DataTableCollection table = result.Tables;
            DataTable resultTable = table[sheetName];

            return resultTable;
        }

        public void ExportToExcel(DataTable resultTable, string fileLocation, string sheetName, string color = "", int numberOfLastRow = 0, int startingCellIndex = 1)
        {
          
            XLWorkbook Workbook = new XLWorkbook(fileLocation);
            IXLWorksheet Worksheet = Workbook.Worksheet(sheetName);

            //Gets the last used row
            if (numberOfLastRow == 0)
            {
                numberOfLastRow  = Worksheet.LastRowUsed().RowNumber();
            }
           
          
            //Defines the starting cell for appeding  (Row , Column)    
            IXLCell CellForNewData = Worksheet.Cell(numberOfLastRow + 1, startingCellIndex);
            if (!color.Equals(""))
            {
                for (int i = 0; i < resultTable.Rows.Count; i++)
                {
                    IXLRow RowForNewData = Worksheet.Row(numberOfLastRow + 1 + i);
                    RowForNewData.Style.Font.FontColor = XLColor.Red;
                }
            }
            //InsertData -  the data from the DataTable without the Column names ; InsertTable - inserts the data with the Column names
            CellForNewData.InsertData(resultTable);

            Workbook.SaveAs(fileLocation);
        }
        public void EditCellValue(string fileLocation, string sheetName, string newCellValue, int row, int column)
        {
            XLWorkbook Workbook = new XLWorkbook(fileLocation);
            IXLWorksheet Worksheet = Workbook.Worksheet(sheetName);
            IXLCell CellForNewData = Worksheet.Cell(row, column);
           
            CellForNewData.Clear();
            CellForNewData.SetValue(newCellValue);
            Workbook.SaveAs(fileLocation);
        }
        public void EditRowColor(string fileLocation, string sheetName, int startIndex, string date)
        {
            XLWorkbook Workbook = new XLWorkbook(fileLocation);
            IXLWorksheet Worksheet = Workbook.Worksheet(sheetName);
            int NumberOfLastRow = Worksheet.LastRowUsed().RowNumber();

            for (int i = startIndex; i < NumberOfLastRow; i++)
            {
                var cellValue = (Worksheet.Cell(i, 4).Value).ToString();
                if (cellValue.Equals(""))
                {
                    break;
                }
                var validity = DateTime.ParseExact(cellValue, "dd.MM.yyyy", null);
                var todayDate = DateTime.ParseExact(date, "dd.MM.yyyy",null);
                IXLRow excelRow = Worksheet.Row(i);
                if (DateTime.Compare(validity, todayDate) < 0)
                {
                    excelRow.Style.Fill.BackgroundColor = XLColor.Red;
                }
                else if (DateTime.Compare(validity, todayDate) == 0)
                {
                    excelRow.Style.Fill.BackgroundColor = XLColor.Yellow;
                }
            }
            Workbook.SaveAs(fileLocation);
        }

    }
  
}
