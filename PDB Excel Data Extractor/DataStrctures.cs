
using System.Data;

namespace PDB_Excel_Data_Extractor
{
    public class DataStrctures
    {
        public DataTable IncomeDataTableStructure()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Date", typeof(string));
            table.Columns.Add("Receipt", typeof(string));
            table.Columns.Add("Studio", typeof(string));
            table.Columns.Add("FirstAndFamilyName", typeof(string));
            table.Columns.Add("CardName", typeof(string));
            table.Columns.Add("AddditionalInfo", typeof(string));
            table.Columns.Add("InstructorName", typeof(string));
            table.Columns.Add("Money", typeof(string));
            return table;
        }
        public DataTable ExpenseDataTableStructure()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Date", typeof(string));
            table.Columns.Add("Studio", typeof(string));
            table.Columns.Add("TypeOfExpense", typeof(string));
            table.Columns.Add("InstructorName", typeof(string));
            table.Columns.Add("Honorary", typeof(string));
            return table;
        }
        public DataTable CardValidityTableStructure()
        {
            DataTable table = new DataTable();
            table.Columns.Add("NumberOfCard", typeof(string));
            table.Columns.Add("FirstAndFamilyName", typeof(string));
            table.Columns.Add("DateFrom", typeof(string));
            table.Columns.Add("DateTo", typeof(string));
            table.Columns.Add("TypeOfCard", typeof(string));
            table.Columns.Add("WayOfPaying", typeof(string));
            table.Columns.Add("ActualAmmount", typeof(int));
            return table;
        }

        public DataTable IndexGetterStructure()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("Index", typeof(int));
            return table;
        }
    }
}
