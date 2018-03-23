using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PDB_Excel_Data_Extractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ExcelPopulationLogic excel = new ExcelPopulationLogic();
            
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
        
            int year = int.Parse(YearInput.Text);
            int month = int.Parse(MonthInput.Text);
            int day = int.Parse(DayInput.Text);
            ExcelPopulationLogic excel = new ExcelPopulationLogic();
            excel.summary(year, month, day);
        }
        private void Button_Click_Expense(object sender, RoutedEventArgs e)
        {
            int year = int.Parse(YearInputFolders.Text);
            int month = int.Parse(MonthInputFolders.Text);
            ExcelPopulationLogic excel = new ExcelPopulationLogic();
            excel.SeedingSharedData(year, month);
        }
    }
}
