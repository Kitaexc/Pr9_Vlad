

using Microsoft.Win32;
using Spire.Xls;
using System.Windows;
using System.Windows.Controls;

namespace Word
{
    /// <summary>
    /// Логика взаимодействия для Redactor4.xaml
    /// </summary>
    public partial class Redactor4 : Window
    {
        public Redactor4()
        {
            InitializeComponent();
        }

        private void Create_Excel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Выберите Excel-файл для работы с ним";

            if (openFileDialog.ShowDialog() == true)
            {

                string filePath = openFileDialog.FileName;
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();
                workbook.LoadFromFile(filePath);
                Worksheet sheet = workbook.Worksheets[0];
                CellRange range = sheet.AllocatedRange;
                var datatable = sheet.ExportDataTable(range, true);
                Table.ItemsSource = datatable.DefaultView;
            }
        }

        private void Send_Excel_Click(object sender, RoutedEventArgs e)
        {
          
        }



    }
}
