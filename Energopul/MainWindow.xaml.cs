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
using System;
using System.IO;
using System.Data;
using System.Data.Common;
using System.Windows;
using System.Windows.Controls;
using System.Data.SQLite;
using System;
using System.Data;
using Path = System.IO.Path;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SQLite;
using System.IO;

namespace Energopul
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadDataToDataGrid();
        }

        private void LoadDataToDataGrid()
        {
            string relativePath = "energopul.db";
            string fullPath = Path.Combine(AppContext.BaseDirectory, relativePath);
            string connectionString = $"Data Source={fullPath};Version=3;";
            using var connection = new SQLiteConnection(connectionString);
            connection.Open();
        
            const string query = "SELECT * FROM Contracts;";
            using var command = new SQLiteCommand(query, connection);
            using var adapter = new SQLiteDataAdapter(command);
            var dataTable = new DataTable();
            adapter.Fill(dataTable);
            Table.ItemsSource = dataTable.DefaultView;
        }
        
        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            string relativePath = "energopul.db";
            string fullPath = Path.Combine(AppContext.BaseDirectory, relativePath);
            string connectionString = $"Data Source={fullPath};Version=3;";
            using var connection = new SQLiteConnection(connectionString);
            connection.Open();
            {
                string query = "SELECT * FROM Contracts";
                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, connection))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    
                    ExportDataTableToExcel(dataTable, "Contracts.xlsx");
                }
            }
        }
        static void ExportDataTableToExcel(DataTable dataTable, string fileName)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Запись заголовков
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                }

                // Запись данных
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                    }
                }

                package.Save();
            }
        }
    }
}
