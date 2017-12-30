using System;
using System.Globalization;
using System.IO;
using System.Data;


namespace PDB_Excel_Data_Extractor
{
    public class Common
    {
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
                // Use static Path methods to extract only the file name from the path.
                fileName = Path.GetFileName(s);
                destFile = Path.Combine(destinationPath, fileName);
                if (FilledFileCheck(s)|| emptyCheck)
                {
                    File.Copy(s, destFile, true);
                }
                else
                {
                    fileName = s.Substring(0, s.Length - 5)+"-EMPTY.xlsx";
                    File.Move(s, fileName);
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

        internal static DataTable LowestOpeningBalance(DataTable table, string studio)
        {
            double balance = 0;
            DataStrctures tables = new DataStrctures();
            DataTable tableToReturn = tables.StartingBalanceTableStructure();
            IndexGetters indexes = new IndexGetters();
            var indexOfRanges = indexes.ListOfAllRanges(table);
            var balanceString = "";
            var time = "";
            for (int i = 0; i < indexOfRanges.Count; i++)
            {
                int range = indexOfRanges[i];
               
                balanceString = table.Rows[range + 4][0].ToString().Substring(3);
                 if (!balanceString.Equals(""))
                {
                    balance = Convert.ToDouble(balanceString);
                    time = table.Rows[range][0].ToString().Substring(0, 5);
                    tableToReturn.Rows.Add(time, balance, studio);
                }
            }
            return tableToReturn;
        }
        public static void CardExpirationValidation(string cardValidityTo)
        {

            var validity = Convert.ToDateTime("21.12.2017");
            var todayDate = Convert.ToDateTime("22.12.2017");
            int ggg = DateTime.Compare(validity, todayDate);

            if (DateTime.Compare(validity, todayDate) < 0)
            {
                Console.WriteLine("Red");
            }
            else if (DateTime.Compare(validity, todayDate) == 0)
            {
                Console.WriteLine("Yellow");

            }
        }


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
