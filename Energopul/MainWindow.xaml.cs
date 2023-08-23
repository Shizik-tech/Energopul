using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using OfficeOpenXml;
using Word = Microsoft.Office.Interop.Word;

namespace Energopul
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadDataToDataGrid();
        }

        private string GetConnectionString()
        {
            string relativePath = "energopul.db";
            string fullPath = Path.Combine(AppContext.BaseDirectory, relativePath);
            return $"Data Source={fullPath};Version=3;";
        }

        private DataTable GetDataFromSqLite(string query)
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                using var command = new SQLiteCommand(query, connection);
                using var adapter = new SQLiteDataAdapter(command);
                var dataTable = new DataTable();
                adapter.Fill(dataTable);
                return dataTable;
            }
        }

        private void LoadDataToDataGrid()
        {
            const string query = "SELECT * FROM Contracts;";
            DataTable dataTable = GetDataFromSqLite(query);
            Table.ItemsSource = dataTable.DefaultView;
        }

        private void ExportDataTableToExcel(DataTable dataTable, string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = null;

                if (package.Workbook.Worksheets.Any(sheet => sheet.Name == "Таблица1"))
                {
                    worksheet = package.Workbook.Worksheets["Таблица1"];
                }
                else
                {
                    worksheet = package.Workbook.Worksheets.Add("Таблица1");
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                    }
                }

                int lastUsedRow = worksheet.Dimension?.End.Row ?? 1;

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[lastUsedRow + row + 1, col + 1].Value = dataTable.Rows[row][col];
                    }
                }

                package.Save();
            }
        }

        
        private void ExportDataTableToWord(DataTable dataTable, string fileName)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();
            
            Word.Paragraph para = doc.Paragraphs.Add();
            Word.Table table = doc.Tables.Add(para.Range, dataTable.Rows.Count + 1, dataTable.Columns.Count);

            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                table.Cell(1, col + 1).Range.Text = dataTable.Columns[col].ColumnName;
            }
            
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    table.Cell(row + 2, col + 1).Range.Text = dataTable.Rows[row][col].ToString();
                }
            }

            doc.SaveAs2(fileName);
            doc.Close();
            wordApp.Quit();

            Marshal.ReleaseComObject(table);
            Marshal.ReleaseComObject(para);
            Marshal.ReleaseComObject(doc);
            Marshal.ReleaseComObject(wordApp);
        }

        private void ExportDataButton_Click(object sender, RoutedEventArgs e)
        {
            string query = "SELECT * FROM Contracts";
            DataTable dataTable = GetDataFromSqLite(query);
            ExportDataTableToExcel(dataTable, "Contracts.xlsx");
            MessageBox.Show("Экспорт данных успешно выполнен", "Экспорт", MessageBoxButton.OK);
        }
        private void ExportDataToWordButton_Click(object sender, RoutedEventArgs e)
        {
            string query = "SELECT * FROM Contracts";
            DataTable dataTable = GetDataFromSqLite(query);
            ExportDataTableToWord(dataTable, "Contracts.docx");
        }
    }
}
