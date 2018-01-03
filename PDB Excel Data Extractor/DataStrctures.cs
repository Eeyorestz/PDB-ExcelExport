
using System;
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
        public DataTable AvailabilityTableStructure(string date, double lowestBalance, double income, double honorary, string clarifications, double cashRegisterValue )
        {
            DataTable table = new DataTable();
            table.Columns.Add("Date", typeof(string));
            table.Columns.Add("TypeOfBalance", typeof(string));
            table.Columns.Add("Sum", typeof(double));
            table.Columns.Add("Clarifications", typeof(string));
            table.Columns.Add("SecondTypeOfBalance", typeof(string));
            table.Columns.Add("Secondum", typeof(double));
         
            table.Rows.Add(Convert.ToDateTime(date).ToString().Substring(0,12), "НС", lowestBalance, clarifications, "НС", lowestBalance);
            table.Rows.Add(Convert.ToDateTime(date).ToString().Substring(0, 12), "Приход", income, clarifications, "Оборот ФУ", cashRegisterValue);
            table.Rows.Add(Convert.ToDateTime(date).ToString().Substring(0, 12), "Разход", honorary, clarifications, "Изгеглени", 0);
            table.Rows.Add(Convert.ToDateTime(date).ToString().Substring(0, 12), "КС", (lowestBalance + income) - honorary, clarifications, "КС");

            return table;
        }
        public DataTable StartingBalanceTableStructure()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Time", typeof(string));
            table.Columns.Add("Balance", typeof(double));
            table.Columns.Add("CashRegister", typeof(double));
            table.Columns.Add("Studio", typeof(string));
            return table;
        }
    }
}
