using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace Word
{
    /// <summary>
    /// Логика взаимодействия для Redactor2.xaml
    /// </summary>
    public partial class Redactor2 : Window
    {
        private DataTable dataTable;
        private Workbook workbook;
        private Worksheet sheet;

        public Redactor2()
        {
            InitializeComponent();
            InitializeExcel();
            InitializeDataGrid();
        }

        private void InitializeExcel()
        {
            workbook = new Workbook();
            sheet = workbook.Worksheets.Add("List 1");

            dataTable = new DataTable();
            dataTable.Columns.Add("Column1");
            dataTable.Columns.Add("Column2");

            Table.ItemsSource = dataTable.DefaultView;
        }

        private void InitializeDataGrid()
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("Column1");
            dataTable.Columns.Add("Column2");
            Table.ItemsSource = dataTable.DefaultView;
        }

        private void Create_Excel_Click(object sender, RoutedEventArgs e)
        {
            if (Table.ItemsSource == null)
            {
                MessageBox.Show("No data to save");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog.Title = "Save Excel-файл";

            try
            {
                if (File.Exists(saveFileDialog.FileName))
                {
                    Workbook workbook = new Workbook();
                    workbook.Worksheets.Clear();
                    Worksheet sheet = workbook.Worksheets[0];
                    var dataview = Table.ItemsSource as DataView;
                    sheet.InsertDataView(dataview, true, 1, 1);
                    workbook.SaveToFile(saveFileDialog.FileName, FileFormat.Version2010);
                }
                else
                {
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        Workbook workbook = new Workbook();
                        workbook.Worksheets.Clear();
                        Worksheet sheet = workbook.Worksheets.Add("New List");
                        var dataview = Table.ItemsSource as DataView;
                        sheet.InsertDataView(dataview, true, 1, 1);
                        workbook.SaveToFile(saveFileDialog.FileName, FileFormat.Version2010);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Save Error: " + ex);
            }
        }

        private void Send_Excel_Click(object sender, RoutedEventArgs e)
        {
        }

        private void Write_Click(object sender, RoutedEventArgs e)
        {
            if (dataTable != null)
            {
                int nextColumnIndex = dataTable.Columns.Count + 1;
                string newColumnName = "Column" + nextColumnIndex;
                dataTable.Columns.Add(newColumnName);

                Table.Columns.Add(new DataGridTextColumn
                {
                    Header = newColumnName,
                    Binding = new Binding(string.Format("[{0}]", newColumnName))
                });
            }
        }
    }
}
