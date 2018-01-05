
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;



namespace PDB_Excel_Data_Extractor
{
    public class ExcelPopulationLogic : Common
    {

        #region InstructorsSheetIndexes
        private int nOfCardIndex ;
        private int typeOfGoodIndex ;
        private int firstAndFamiliyNameIndex ;
        private int honoraryIndex ;
        private int moneyIndex ;
        private int receiptIndex ;
        private int additionalInfoIndex;


        #endregion
        #region CardInfoIndexes
        private int poleDanceIndex;
        private int stretchingIndex;
        private int hathaYogaIndex;
        private int airYogaIndex;
        private int classicYogaIndex;
        private int aerialPoleIndex;
        private int exocitPoleDanceIndex;
        private int kidsYogaIndex;
        private int handStandIndex;
        private int aerialYogaKids;
        #endregion
        private string date = "";
        private double CenterHonorary = 0;
        private double CenteryIncome = 0;
        private double LozenecHonorary = 0;
        private double LozenecIncome = 0;
        private double StudentskiHonorary = 0;
        private double StudentskiIncome = 0;

        private List<DataTable> StartingBalances = new List<DataTable>();

        readonly string  expenseFile = AssemblyDirectory + @"\Zimnina\_Лютеница.xlsx";
        public void SeedingSharedData(int year, int month)
        {
            FolderPopulation folders = new FolderPopulation();            
            if (MessageBox.Show("Искате ли да продължите ?",
                "Confirmation", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                folders.PopulateMontlyFiles(year, month);
            }
            MessageBox.Show("Завършена операция",
               "Confirmation", MessageBoxButton.OK);
        }
        public void summary(int year, int month, int day)
        {
             FolderPopulation folders = new FolderPopulation();
             folders.ExtractDataToArchive(year, month, day);
             PopulatingForInstructors(year, month, day);
        }

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
                    
                    ColumnIndexGetterInstructorFile(sheetInfo);
                  
                    string studioName = Path.GetFileName(namesOfStudios[p]);
                    studioName = studioName.Substring(0, studioName.Length - 5);
                   

                    if (studioName.Equals("Студентски"))
                    {
                        reader.ExportToExcel(IncomeDataTable(sheetInfo, Instructors()[g], studioName, ТrueExpense(sheetInfo)), AssemblyDirectory + @"\Zimnina\_Компот-Студентски.xlsx", "Приход");
                        reader.ExportToExcel(ExpenseDataTable(sheetInfo, Instructors()[g], studioName, year, month, day), AssemblyDirectory + @"\Zimnina\_Компот-Студентски.xlsx", "Разход");
                    }
                    else
                    {
                        reader.ExportToExcel(IncomeDataTable(sheetInfo, Instructors()[g], studioName, ТrueExpense(sheetInfo)), AssemblyDirectory + @"\Zimnina\_Компот-ЦЛ.xlsx", "Приход");
                        reader.ExportToExcel(ExpenseDataTable(sheetInfo, Instructors()[g], studioName, year, month, day), AssemblyDirectory + @"\Zimnina\_Компот-ЦЛ.xlsx", "Разход");
                    }

                    List<DataTable> data = CardValidityDataTable(sheetInfo, ТrueExpense(sheetInfo), year, month, day);
                    for (int w = 0; w < data.Count; w++)
                    {
                        if (data[w].Rows[0]["WayOfPaying"].ToString().Equals("50%"))
                        {
                            reader.ExportToExcel(data[w], expenseFile, "Справка карти", "Red");
                        }
                        else
                        {
                            reader.ExportToExcel(data[w], expenseFile, "Справка карти");
                        }
                    }
                    StartingBalances.Add(LowestOpeningBalance(sheetInfo, studioName));
                    //CardValidityPupulation(sheetInfo);
                }

            }
            AvailabilityDataTable(sheetInfo);
            MessageBox.Show("Завършена операция",
                "Confirmation", MessageBoxButton.OK);
            {
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
                switch (studioName)
                {
                    case "Лозенец":
                        LozenecIncome = LozenecIncome+Convert.ToDouble(Money);
                        break;
                    case "Студентски":
                        StudentskiIncome = StudentskiIncome + Convert.ToDouble(Money);
                        break;
                    case "Център":
                        CenteryIncome = CenteryIncome + Convert.ToDouble(Money);
                        break;
                }
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
                    honorarySumToAdd = DelimterConvertor(honorary);
                }
                honorarySum = honorarySum + honorarySumToAdd;
            }
            switch (studioName)
            {
                case "Лозенец":
                    LozenecHonorary = LozenecHonorary+ honorarySum;
                    break;
                case "Студентски":
                    StudentskiHonorary = StudentskiHonorary + honorarySum;
                    break;
                case "Център":
                    CenterHonorary = CenterHonorary + honorarySum;
                    break;
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
                if (!cardName.Equals("")|| typeOfood.Equals("ваучер"))
                {
                    table.Rows.Add(cardName, firstAndFamilyName, date, ValidityTo, typeOfood ,wayOfPaying, ammountOfMoney);
                    listOfTables.Add(table);
                }
            }
            return listOfTables;
        }

        private void AvailabilityDataTable(DataTable sheetInfo)
        {
            DataStrctures structure = new DataStrctures();
          
            ExcelReader reader = new ExcelReader();
            DataTable BalanceSheet = reader.ExcelToDataTable("Наличност", AssemblyDirectory + @"\Zimnina\_Компот-ЦЛ.xlsx" );
            int startingRow =   StartingRowForExport(BalanceSheet, date);

         
          
            for (int i = 0; i < StudioNames().Count; i++)
            {
                var lowestBalance = LowestAmmountPopulating(StudioNames()[i], StartingBalances,"Balance");
                var cashRegisterLowestBalance = LowestAmmountPopulating(StudioNames()[i], StartingBalances, "CashRegister");
                DataTable table = new DataTable();
               
                if (StudioNames()[i].Equals("Студентски"))
                {
                    table = structure.AvailabilityTableStructure(date, lowestBalance, StudentskiIncome, StudentskiHonorary, "", cashRegisterLowestBalance);
                    reader.ExportToExcel(table, AssemblyDirectory + @"\Zimnina\_Компот-Студентски.xlsx", "Наличност", numberOfLastRow: startingRow);
                }
                else
                {
                    if (StudioNames()[i].Equals("Лозенец"))
                    {
                        table = structure.AvailabilityTableStructure(date, lowestBalance, LozenecIncome, LozenecHonorary, "", cashRegisterLowestBalance);
                        reader.ExportToExcel(table, AssemblyDirectory + @"\Zimnina\_Компот-ЦЛ.xlsx", "Наличност", numberOfLastRow: startingRow, startingCellIndex: 8);
                    }
                    else
                    {
                        table = structure.AvailabilityTableStructure(date, lowestBalance, CenteryIncome, CenterHonorary, "", cashRegisterLowestBalance);
                        reader.ExportToExcel(table, AssemblyDirectory + @"\Zimnina\_Компот-ЦЛ.xlsx", "Наличност", numberOfLastRow: startingRow);
                    }
                }
            }
          
        }

    
        private void CardValidityPupulation(DataTable sheetInfo)
        {
            IndexGetters indexes = new IndexGetters();
            ExcelReader excel = new ExcelReader();
            DataTable expenseInfo = excel.ExcelToDataTable("Справка карти", expenseFile);
            var workouts = indexes.ListOfWorkouts(sheetInfo);
            var workoutsIndexes = indexes.Ranges(sheetInfo);
            var workoutType = "";
            
            for (int i = 0; i < sheetInfo.Rows.Count; i++)
            {
                var cardName = sheetInfo.Rows[i][nOfCardIndex].ToString();
                var typeOfGood = sheetInfo.Rows[i][typeOfGoodIndex].ToString();
                var ammount = 0;
                if (!cardName.Equals("") && !cardName.Equals("N: на карта")&& !typeOfGood.Equals("ваучер"))
                {
                    for (int j = 0; j < workouts.Count; j++)
                    {
                        if (i >= workoutsIndexes[j][0] && i < workoutsIndexes[j][1])
                        {
                            workoutType = workouts[j];
                            var cardRowNumber = indexes.RowOfCard(expenseInfo, cardName);
                            var cardColumnNumber = ColumnIndexGetterCardInfoFile(expenseInfo, workoutType);
                            if (!expenseInfo.Rows[cardRowNumber][cardColumnNumber].ToString().Equals(""))
                            {
                                ammount = int.Parse(expenseInfo.Rows[cardRowNumber][cardColumnNumber].ToString());
                            }
                            var sum = SumForWorkoutType(ammount, workoutType).ToString();
                            excel.EditCellValue(expenseFile, "Справка карти", sum ,cardRowNumber+1, cardColumnNumber+1);                       
                            break;
                        }
                    }
                }
            }
            excel.EditRowColor(expenseFile, "Справка карти", (indexes.CardExpirationStartingRowIndex(expenseInfo) + 2), date);
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

        private List<string> StudioNames()
        {
            List<string> list = new List<string>();
            list.Add("Студентски");
            list.Add("Лозенец");
            list.Add("Център");
            return list;
        }

        #region IndexPopulation
        private void ColumnIndexGetterInstructorFile(DataTable sheetInfo)
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
        private int ColumnIndexGetterCardInfoFile(DataTable sheetInfo, string typeOfWorkout)
        {
            IndexGetters index = new IndexGetters();
            int returnIndex = 0;
            for (int i = 0; i < sheetInfo.Columns.Count; i++)
            {
                if (sheetInfo.Rows[index.CardExpirationStartingRowIndex(sheetInfo)][i].ToString().ToLower().Equals(typeOfWorkout))
                {
                    returnIndex = i;
                }
            }
            return returnIndex;
        }
        #endregion

    }
}
