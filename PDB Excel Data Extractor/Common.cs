using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDB_Excel_Data_Extractor
{
    public class Common
    {
        internal static string[] Instructors()
        {
            string[] lines = File.ReadAllLines(@"C:\PDB\Instructors.txt");
            return lines;
        }

        internal static string MonthName(int month)
        {
            string monthName = CultureInfo.CurrentUICulture.DateTimeFormat.GetMonthName(month);
            return monthName;
        }

        internal static string InstructorsDirectory(int year, string instructorName)
        {
            string instructorsDirectory = @"C:\PDB\Share\" + year + "" + @"\" + instructorName + "";
            return instructorsDirectory;
        }

        internal static string ArchiveDayDirectory(int year, string monthName, int day, string instructorName)
        {
            string dayDirectory = @"C:\PDB\Archive\" + year + "" + @"\" + monthName + @"\" + day + @"\" + instructorName;
            return dayDirectory;
        }

        internal static void CopyPasteFiles(string sourcePath, string destinationPath)
        {
            string fileName = "";
            string destFile = "";
            string[] files = Directory.GetFiles(sourcePath);

            foreach (string s in files)
            {
                // Use static Path methods to extract only the file name from the path.
                fileName = Path.GetFileName(s);
                destFile = Path.Combine(destinationPath, fileName);
                File.Copy(s, destFile, true);
            }
        }

        internal static string DateOfExportedFile(int year, int month, int day)
        {
            var dat = new DateTime(year, month, day);
            string date = dat.ToString("dd.MM.yyyy");
            return date;
        }

        internal static string[] GetFileNames(int year, string monthName, int day, string instructorName)
        {
            string[] files = Directory.GetFiles(ArchiveDayDirectory(year, monthName, day, instructorName));
            return files;
        }

        internal static double delimterConvertor(string number)
        {
            if (number.Contains("."))
            {
                number = number.Replace(".", ",");
            }
            return Double.Parse(number);
        }
    }
}
