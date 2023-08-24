using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using Word = Microsoft.Office.Interop.Word;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XWPF.UserModel;

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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {
                MessageBoxResult result = MessageBox.Show("Документ с таким названием уже существует. Перезаписать?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    MessageBox.Show("Операция отменена.");
                    return;
                }
                else
                {
                    fileInfo.Delete();
                }
            }

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

            MessageBox.Show("Экспорт данных успешно выполнен", "Экспорт", MessageBoxButton.OK);
        }

        private void ExportDataTableToWord(DataTable dataTable, string fileName)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());


                Table table = new Table();


                TableProperties tblProperties = new TableProperties(
                    new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "100%" });

                table.AppendChild(tblProperties);


                TableRow headerRow = new TableRow();

                foreach (DataColumn column in dataTable.Columns)
                {
                    TableCell headerCell = new TableCell(new Paragraph(new Run(new Text(column.ColumnName))));
                    headerRow.Append(headerCell);
                }

                table.Append(headerRow);


                foreach (DataRow dataRow in dataTable.Rows)
                {
                    TableRow dataTableRow = new TableRow();

                    foreach (var item in dataRow.ItemArray)
                    {
                        TableCell dataCell = new TableCell(new Paragraph(new Run(new Text(item.ToString()))));
                        dataTableRow.Append(dataCell);
                    }

                    table.Append(dataTableRow);
                }

                body.Append(table);
            }
        }
        
        private void ExportDataButton_Click(object sender, RoutedEventArgs e)
        {
            string query = "SELECT * FROM Contracts";
            DataTable dataTable = GetDataFromSqLite(query);
            ExportDataTableToExcel(dataTable, "Contracts.xlsx");
        }
        
        private void ExportDataToWordButton_Click(object sender, RoutedEventArgs e)
        {
            string query = "SELECT * FROM Contracts";
            DataTable dataTable = GetDataFromSqLite(query);
            ExportDataTableToWord(dataTable, "Contracts.docx");
        }

        private void ButtonShow_Click(object sender, EventArgs e)
        {
            int periodIndex = Period.SelectedIndex + 1; // Индексы ComboBox начинаются с 0
            var conn = new SQLiteConnection(GetConnectionString());
            using (SQLiteCommand cmd = conn.CreateCommand())
            {
                string end_date_rout = DateTime.Now.ToString("yyyy-MM-dd");

                string start_date = "";
                switch (periodIndex)
                {
                    case 1:
                        start_date = DateTime.Now.AddMonths(-3).ToString("yyyy-MM-dd");
                        break;
                    case 2:
                        start_date = DateTime.Now.AddMonths(-6).ToString("yyyy-MM-dd");
                        break;
                    case 3:
                        start_date = DateTime.Now.AddYears(-1).ToString("yyyy-MM-dd");
                        break;
                    case 4:
                        start_date = DateTime.Now.AddYears(-2).ToString("yyyy-MM-dd");
                        break;
                    default:
                        MessageBox.Show("Invalid period");
                        return;
                }

                cmd.CommandText = $@"
                    SELECT
                        org.org_name AS 'Название организации',
                        org.INN_org AS 'ИНН организации',
                        con.сon_num AS 'Номер договора',
                        con.date_of_con AS 'Дата заключения договора',
                        con.end_date AS 'Дата окончания договора',
                        con.sub_of_contr AS 'Предмет договора',
                        con.con_sum AS 'Сумма договора',
                        con.deadlin_stages AS 'Сроки по этапам договора'
                    FROM
                        Contracts con
                        JOIN Organizations org ON con.org_id = org.id
                    WHERE
                        con.date_of_con BETWEEN @start_date AND @end_date_rout";

                cmd.Parameters.AddWithValue("@start_date", start_date);
                cmd.Parameters.AddWithValue("@end_date_rout", end_date_rout);

                DataTable dataTable = new DataTable();
                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd))
                {
                    adapter.Fill(dataTable);
                }

                // Привязка данных к WPF DataGrid (Table)
                Table.ItemsSource = dataTable.DefaultView;
            }
        }



        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }
    }
}

