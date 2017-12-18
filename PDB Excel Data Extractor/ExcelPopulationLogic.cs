
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;



namespace PDB_Excel_Data_Extractor
{
    public class ExcelPopulationLogic : Common
    {
        private int nOfCardIndex = 0;
        private int typeOfGoodIndex = 0;
        private int firstAndFamiliyNameIndex = 0;
        private int honoraryIndex = 0;
        private int moneyIndex = 0;
        private int receiptIndex = 0;
        private int additionalInfoIndex = 0;
        private string date = "";

        public void SeedingSharedData(int year, int month)
        {
            FolderPopulation folders = new FolderPopulation();            
            if (MessageBox.Show("Искате ли да продължите ?",
                "Confirmation", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                folders.PopulateMontlyFiles(year, month);
            }
        }
        public void summary(int year, int month, int day)
        {
             FolderPopulation folders = new FolderPopulation();
            // folders.ExtractDataToArchive(year, month, day);
             PopulatingForInstructors(year, month, day);
        }

        // Use static Path methods to extract only the file name from the path.
        private void PopulatingForInstructors(int year, int month, int day)
        {
            ExcelReader reader = new ExcelReader();
            DataTable sheetInfo = null;
            date = DateOfExportedFile(year, month, day);
            for (int g = 0; g < Instructors().Length; g++)
            {
                string[] namesOfStudios = GetFileNames(year, MonthName(month), day, Instructors()[g]);
                for (int p = 0; p < namesOfStudios.Length; p++)
                {
                    sheetInfo = reader.ExcelToDataTable("Sheet1", namesOfStudios[p]);
                    string studioName = Path.GetFileName(namesOfStudios[p]);
                    studioName = studioName.Substring(0, studioName.Length - 5);
                    ColumnIndexGetter(sheetInfo);
                   // reader.ExportToExcel(ExpenseDataTable(sheetInfo, Instructors()[g], studioName, year, month, day), @"C:\PDB\_Компот-ЦЛ-Декември.xlsx", "Разход");
                   // reader.ExportToExcel(IncomeDataTable(sheetInfo, Instructors()[g], studioName, ТrueExpense(sheetInfo)), @"C:\PDB\_Компот-ЦЛ-Декември.xlsx", "Приход");
                    List<DataTable> data = CardValidityDataTable(sheetInfo, ТrueExpense(sheetInfo), year, month, day);
                    for (int w = 0; w < data.Count; w++)
                    {
                        
                      
                        if (data[g].Rows[0]["WayOfPaying"].ToString().Equals("50%"))
                        {
                            reader.ExportToExcel(data[g], @"C:\PDB\_Компот-ЦЛ-Декември.xlsx", "Приход", "Red");
                        }
                        else
                        {
                            reader.ExportToExcel(data[g], @"C:\PDB\_Компот-ЦЛ-Декември.xlsx", "Приход");
                        }
                    }
                }
            }
        }

        private DataTable IncomeDataTable(DataTable sheetInfo,string instrctorName, string studioName ,List<int> trueExpenses)
        {
            DataStrctures structure = new DataStrctures();
            DataTable table = structure.IncomeDataTableStructure();

            for (int i = 0; i < trueExpenses.Count; i++)
            {
                int indexOfRow = trueExpenses[i];
                var typeOfGood = sheetInfo.Rows[indexOfRow][typeOfGoodIndex].ToString();
                var FirstAndFamilyName = sheetInfo.Rows[indexOfRow][firstAndFamiliyNameIndex].ToString();
                var Money = sheetInfo.Rows[indexOfRow][moneyIndex].ToString();
                var Receipt = sheetInfo.Rows[indexOfRow][receiptIndex].ToString();
                if (!Receipt.ToLower().Equals("да"))
                {
                    Receipt = "-";
                }
                var AddditionalInfo = sheetInfo.Rows[indexOfRow][additionalInfoIndex].ToString();
                table.Rows.Add(date, Receipt, studioName, FirstAndFamilyName, typeOfGood, AddditionalInfo, instrctorName, Money);
            }
            return table;
        }
        private DataTable ExpenseDataTable(DataTable sheetInfo, string instrctorName, string studioName,  int year, int month, int day)
        {
            DataStrctures structure = new DataStrctures();
            DataTable table = structure.ExpenseDataTableStructure();
            string motive = "инструкторски хонорар";
            double honorarySum = 0;
            double honorarySumToAdd = 0;
            for (int i = 0; i < sheetInfo.Rows.Count; i++)
            {
                string honorary = sheetInfo.Rows[i][honoraryIndex].ToString();
                if (honorary.Equals("")|| honorary.Equals("хонорар"))
                {
                    honorarySumToAdd = 0;
                }
                else
                {
                    honorarySumToAdd = delimterConvertor(honorary);
                }
                honorarySum = honorarySum + honorarySumToAdd;
            }
            table.Rows.Add(date, studioName, motive, instrctorName, honorarySum);
            return table;
        }
        private List<DataTable> CardValidityDataTable(DataTable sheetInfo, List<int> trueExpenses, int year, int month, int day)
        {
            DataStrctures structure = new DataStrctures();
           
            List<DataTable> listOfTables = new List<DataTable>();
            MoneyData moneyData = new MoneyData();
           
            for (int i = 0; i < trueExpenses.Count; i++)
            {
                DataTable table = structure.CardValidityTableStructure();
                int indexOfRow = trueExpenses[i];
                var money = sheetInfo.Rows[indexOfRow][moneyIndex].ToString();
                var cardName = sheetInfo.Rows[indexOfRow][nOfCardIndex].ToString();
                var firstAndFamilyName = sheetInfo.Rows[indexOfRow][firstAndFamiliyNameIndex].ToString();
                var wayOfPaying = moneyData.deferredPayment(money);

                 var typeOfood = sheetInfo.Rows[indexOfRow][typeOfGoodIndex].ToString();
                var ammountOfMoney = moneyData.Ammout(money,
                   typeOfood);
                var ValidityTo = moneyData.CardPeriodExpiration(year, month, day, ammountOfMoney);
                if (!cardName.Equals(""))
                {
                    table.Rows.Add(cardName, firstAndFamilyName, date, ValidityTo, typeOfood ,wayOfPaying, ammountOfMoney);
                    listOfTables.Add(table);
                }
            }
            return listOfTables;
        }
        private List<int> ТrueExpense(DataTable sheetInfo)
        {
            List<int> list = new List<int>();
            for (int i = 0; i < sheetInfo.Rows.Count; i++)
            { 
                if (!sheetInfo.Rows[i][moneyIndex].ToString().Equals("") &&
                     !sheetInfo.Rows[i][moneyIndex].ToString().Equals("карта/ пари") 
                    && !sheetInfo.Rows[i][0].ToString().Equals("Общо: "))
                {
                    list.Add(i);
                }
            }
            return list;
        }
        private void ColumnIndexGetter(DataTable sheetInfo)
        {
            for (int i = 0; i < sheetInfo.Columns.Count; i++)
            {
                switch (sheetInfo.Rows[1][i].ToString().ToLower())
                {
                    case "n: на карта":
                        nOfCardIndex = i;
                        break;
                    case "вид карта/стока":
                        typeOfGoodIndex = i;
                        break;
                    case "име и фамилия:":
                        firstAndFamiliyNameIndex = i;
                        break;
                    case "хонорар":
                        honoraryIndex = i;
                        break;
                    case "карта/ пари":
                        moneyIndex = i;
                        break;
                    case "касов бон":
                        receiptIndex = i;
                        break;
                    case "начин на плащане":
                        additionalInfoIndex = i;
                        break;
                }
            }
           
        } 
    }
}
