using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Data;
using System.Linq;


namespace PDB_Excel_Data_Extractor
{
    public class Common
    {
        public static int CardExpirationValidation(string firstDate, string secondDate)
        {
            var firstTime = Convert.ToDateTime(firstDate);
            var secondTime = Convert.ToDateTime(secondDate);
            int result = DateTime.Compare(firstTime, secondTime);
            return result;
        }
        internal static string[] Instructors()
        {
            string[] lines = File.ReadAllLines(AssemblyDirectory+@"\Templates\Instructors.txt");
            return lines;
        }
        internal static string MonthName(int month)
        {
            string monthName = CultureInfo.CurrentUICulture.DateTimeFormat.GetMonthName(month);
            return monthName;
        }
        internal static string InstructorsDirectory(int year, string instructorName)
        {
            string instructorsDirectory = AssemblyDirectory +@"\Share\" + year + "" + @"\" + instructorName + "";
            return instructorsDirectory;
        }
        internal static string ArchiveDayDirectory(int year, string monthName, int day, string instructorName)
        {
            string dayDirectory = AssemblyDirectory+@"\Archive\" + year + "" + @"\" + monthName + @"\" + day + @"\" + instructorName;
            return dayDirectory;
        }
        internal static void CopyPasteFiles(string sourcePath, string destinationPath, bool emptyCheck = false)
        {
            string fileName = "";
            string destFile = "";
            string[] files = Directory.GetFiles(sourcePath);

            foreach (string s in files)
            {
                if (!s.Contains("gsheet") && !s.Contains("ini"))
                {
                    fileName = Path.GetFileName(s);
                    destFile = Path.Combine(destinationPath, fileName);
                    if (FilledFileCheck(s) || emptyCheck)
                    {
                        File.Copy(s, destFile, true);
                    }
                    else
                    {
                        if (!fileName.Contains("EMPTY"))
                        {
                            fileName = s.Substring(0, s.Length - 5) + "-EMPTY.xlsx";
                            File.Move(s, fileName);
                        }
                    }
                }
            }
        }
        internal static string DateOfExportedFile(int year, int month, int day)
        {
            var dat = new DateTime(year, month, day).ToString("dd.MM.yyyy");
            return dat;
        }
        internal static string[] GetFileNames(int year, string monthName, int day, string instructorName)
        {
            string[] files = Directory.GetFiles(ArchiveDayDirectory(year, monthName, day, instructorName));
            return files;
        }
        internal static double DelimterConvertor(string number)
        {
            if (number.Contains("."))
            {
                number = number.Replace(".", ",");
            }
            return Double.Parse(number);
        }
        internal static int SumForWorkoutType(int ammount, string workoutType)
        {
            int sumToAdd = 0;
            
            switch (workoutType)
            {
                case "пол денс":
                case "пол фит":
                case "екзотик пол денс":
                case "въздушна акробатика":
                    sumToAdd = 18;
                    break;
                case "въздушна йога":
                    sumToAdd = 15;
                    break;
                case "стречинг":
                case "йога":
                case "детска йога":
                case "детска акробатика":
                case "хендстенд":
                    sumToAdd = 12;
                    break;
            }
            int sum = ammount + sumToAdd;
            return sum;
        }
        internal static string AssemblyDirectory
        {
            get
            {
                string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        #region BalanceMethods
        internal static DataTable LowestOpeningBalance(DataTable table, string studio)
        {
            double balance = 0;
            double casRegisterBalance = 0;
            DataStrctures tables = new DataStrctures();
            DataTable tableToReturn = tables.StartingBalanceTableStructure();
            IndexGetters indexes = new IndexGetters();
            var indexOfRanges = indexes.ListOfAllRanges(table);
            var balanceString = "";
            var cashRegisterString = "";
            var time = "";
            for (int i = 0; i < indexOfRanges.Count; i++)
            {
                int range = indexOfRanges[i];
              
                balanceString = table.Rows[range + 4][0].ToString().Substring(3);
                cashRegisterString = table.Rows[range + 6][0].ToString().Substring(6);
                if (!balanceString.Equals(""))
                {
                    balance = DelimterConvertor(balanceString);
                    casRegisterBalance = DelimterConvertor(cashRegisterString);
                    time = table.Rows[range][0].ToString().Substring(0, 5);
                    tableToReturn.Rows.Add(time, balance, casRegisterBalance, studio);
                }
            }
            return tableToReturn;
        }
        internal static int DateTimeComparer(string timePeriodOne, string timePeriodTwo)
        {
            int result = DateTime.Compare(Convert.ToDateTime(timePeriodOne), Convert.ToDateTime(timePeriodTwo));
            return result;
        }
        internal static double LowestAmmountPopulating(string studioName, List<DataTable> StartingBalances, string columnName)
        {
            double lowestBalance = 0;
            var studentsTownList = StartingBalances.FindAll(x => x.Rows[0]["Studio"].ToString().Equals(studioName)).ToList();
            if (studentsTownList.Count == 1)
            {
                var balance = DelimterConvertor(studentsTownList[0].Rows[0][columnName].ToString());
                    lowestBalance = balance;
            }
            else
            {
                for (int i = 0; i < studentsTownList.Count - 1; i++)
                {
                    var tempTime = studentsTownList[i].Rows[0]["Time"].ToString();
                    var balance = DelimterConvertor(studentsTownList[i].Rows[0][columnName].ToString());
                    var time = studentsTownList[i + 1].Rows[0]["Time"].ToString();
                    if (lowestBalance == 0)
                    {
                        lowestBalance = balance;
                    }
                    if (DateTimeComparer(tempTime, time) < 0 && lowestBalance >= balance)
                    {
                        lowestBalance = balance;
                    }
                }
            }

            
            return lowestBalance;
        }

        internal static int StartingRowForExport(DataTable table, string date)
        {
            int row = 0;
            for (int i = 2; i < table.Rows.Count; i++)
            {
                var dateTime = Convert.ToDateTime(date).ToString();
                if (DateTimeComparer(table.Rows[i][0].ToString(), dateTime) ==0)
                {
                    row = i;
                    break;
                }
            }
            return row;
        }
      

        #endregion






        private static bool FilledFileCheck(string pathFile)
        {
            bool check = false;
            ExcelReader excel = new ExcelReader();
            DataTable table = excel.ExcelToDataTable("Sheet1",pathFile);
            int nameAndFamilyIndex = 0;
            for (int t = 0; t < table.Columns.Count; t++)
            {
                if (table.Rows[1][t].ToString().Contains("име и фамилия"))
                {
                    nameAndFamilyIndex = t;
                    break;
                }
            }
            for (int i = 1; i < table.Rows.Count; i++)
            {
                if (!table.Rows[i][nameAndFamilyIndex].ToString().Equals("")&& !table.Rows[i][nameAndFamilyIndex].ToString().Contains("име и фамилия"))
                {
                    check = true;
                    break;
                }
            }  
            return check;
        }
        
    }
}
