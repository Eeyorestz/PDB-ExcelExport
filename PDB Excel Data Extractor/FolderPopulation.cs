using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace PDB_Excel_Data_Extractor
{
    public class FolderPopulation : Common
    {
        public  void ExtractDataToArchive(int year, int month, int day)
        {
            for (int i = 0; i < Instructors().Length; i++)
            {
                string dayDirectory = ArchiveDayDirectory(year, MonthName(month), day, Instructors()[i]);
                Directory.CreateDirectory(dayDirectory);
                string sourcePath = InstructorsDirectory(year, Instructors()[i]) + @"\" +
                             day;
                CopyPasteFiles(sourcePath, dayDirectory);
            }
        }
        public  void PopulateMontlyFiles(int year, int month)
        {
            for (int t = 0; t < Instructors().Length; t++)
            {
                Directory.CreateDirectory(InstructorsDirectory(year, Instructors()[t]));
                DeleteDirectories(InstructorsDirectory(year, Instructors()[t]));
                for (int i = 0; i < GetDaysInMonth(year, month); i++)
                {
                    string dayDirectory = InstructorsDirectory(year, Instructors()[t]) + @"\" + (i + 1);
                    Directory.CreateDirectory(dayDirectory);
                    string sourcePath = AssemblyDirectory+@"\Templates\" + DayOfTheWeek(year, month, (i + 1));
                    CopyPasteFiles(sourcePath, dayDirectory, true);
                }
            }
        }
        private  int GetDaysInMonth(int year, int month)
        {
            int daysInMonth = DateTime.DaysInMonth(year, month); ;
            return daysInMonth;
        }
        private  void DeleteDirectories(string directoryPath)
        {
            foreach (var d in Directory.GetDirectories(directoryPath))
            {
                string[] files = Directory.GetFiles(d);
                foreach (string s in files)
                {
                    File.Delete(s);
                }
                Directory.Delete(d);
            }
        }
        private  string DayOfTheWeek(int year, int month, int day)
        {
            DateTime dt = new DateTime(year, month, day);
            string dayOfTheWeek = dt.DayOfWeek.ToString();
            return dayOfTheWeek;
        }    
    }
}
