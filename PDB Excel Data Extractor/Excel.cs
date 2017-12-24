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

        static List<DataCollection> dataColection = new List<DataCollection>();

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

        /// <summary>
        /// Populate the data from an Excel file to Data Collection
        /// </summary>
        /// <param name="fileName">The name of the Excel file</param>
        /// <param name="sheetName">The name of the worksheet in the Excel file</param>
        public void PopulateToDataTable(string sheetName, string filePath)
        {
            DataTable table = ExcelToDataTable( sheetName, filePath);

            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    DataCollection dtTable = new DataCollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    dataColection.Add(dtTable);
                }

            }
        }

        public void ExportToExcel(DataTable resultTable, string fileLocation, string sheetName, string color = "")
        {
            XLWorkbook Workbook = new XLWorkbook(fileLocation);
            IXLWorksheet Worksheet = Workbook.Worksheet(sheetName);

            //Gets the last used row
            int NumberOfLastRow = Worksheet.LastRowUsed().RowNumber();
          
            //Defines the starting cell for appeding  (Row , Column)    
            IXLCell CellForNewData = Worksheet.Cell(NumberOfLastRow + 1, 1);
            if (!color.Equals(""))
            {
                for (int i = 0; i < resultTable.Rows.Count; i++)
                {
                    IXLRow RowForNewData = Worksheet.Row(NumberOfLastRow + 1 + i);
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
            IXLRow Row = Worksheet.Row(row);

            CellForNewData.Clear();
            CellForNewData.SetValue(newCellValue);
            Workbook.SaveAs(fileLocation);
        }
        public void EditRowColor(string fileLocation, string sheetName, int row)
        {
            XLWorkbook Workbook = new XLWorkbook(fileLocation);
            IXLWorksheet Worksheet = Workbook.Worksheet(sheetName);
            IXLRow excelRow = Worksheet.Row(row);
            excelRow.Style.Fill.BackgroundColor = XLColor.Yellow;
            Workbook.SaveAs(fileLocation);
        }
        /// <summary>
        /// Read the data from the a Collection
        /// </summary>
        /// <param name="rowNubmer">The row nubmer</param>
        /// <param name="columnName">The name of the column</param>
        /// <returns></returns>
        public static string ReadData(int rowNubmer, string columnName)
        {
            try
            {
                string data = (from ColData in dataColection
                               where ColData.colName == columnName && ColData.rowNumber == rowNubmer
                               select ColData.colValue).SingleOrDefault();
                return data.ToString();

            }
            catch (Exception e)
            {
                return null;
            }
        }   
    }
  


    /// <summary>
    /// Custom data collection with row number, column name and column value
    /// </summary>
    public class DataCollection
    {
        public int rowNumber { get; set; }
        public string colName { get; set; }
        public string colValue { get; set; }
    }
}
