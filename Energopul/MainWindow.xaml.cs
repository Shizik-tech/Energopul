using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                }

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
                        org.Name AS 'Название организации',
                        org.INN AS 'ИНН организации',
                        con.Con_date AS 'Номер договора',
                        con.Start_date AS 'Дата заключения договора',
                        con.Con_stage AS 'Этап договора',
                        con.Sub_of_con AS 'Предмет договора',
                        con.Con_sum AS 'Сумма договора',
                        con.Con_end AS 'Дата окончания договора'
                    FROM
                        Contracts con
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

        private void SaveChangesToDatabase(DataTable dataTable)
        {
            using var connection = new SQLiteConnection(GetConnectionString());
            connection.Open();

            using (var transaction = connection.BeginTransaction())
            using (var command = new SQLiteCommand(connection))
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    string commandText = "";
                    command.Parameters.Clear();

                    
                    if (row.RowState == DataRowState.Added)
                    {
                        // Операция INSERT
                        commandText = "INSERT INTO Contracts (Название, ИНН, \"Номер договора\", \"Дата заключения\", \"Дата начала выполнения\", \"Предмет договора\", \"Сумма договора\", \"Этап договора\", \"Дата окончания выполнения\") VALUES (@Название, @ИНН, @Номер_договора, @Дата_заключения, @Дата_начала_выполнения, @Предмет_договора, @Сумма_договора, @Этап_договора, @Дата_окончания_выполнения)";
                    }
                    else if (row.RowState == DataRowState.Modified)
                    {
                        // Операция UPDATE
                        commandText = "UPDATE Contracts SET Название = @Название, ИНН = @ИНН, \"Номер договора\" = @Номер_договора, \"Дата заключения\" = @Дата_заключения, \"Дата начала выполнения\" = @Дата_начала_выполнения, \"Предмет договора\" = @Предмет_договора, \"Сумма договора\" = @Сумма_договора, \"Этап договора\" = @Этап_договора, \"Дата окончания выполнения\" = @Дата_окончания_выполнения WHERE id = @id";
                        command.Parameters.AddWithValue("@id", row["id"]);
                    }
                    else if (row.RowState == DataRowState.Deleted)
                    {
                        // Операция DELETE
                        DataRow originalRow = row.Table.Rows.Find(row["id", DataRowVersion.Original]);
                        if (originalRow != null)
                        {
                            commandText = "DELETE FROM Contracts WHERE id = @id";
                            command.Parameters.AddWithValue("@id", originalRow["id"]);
                        }
                    }

                    if (!string.IsNullOrEmpty(commandText))
                    {
                        command.CommandText = commandText;
                        command.Parameters.AddWithValue("@Название", row["Название"]);
                        command.Parameters.AddWithValue("@ИНН", row["ИНН"]);
                        command.Parameters.AddWithValue("@Номер_договора", row["Номер договора"]);
                        command.Parameters.AddWithValue("@Дата_заключения", row["Дата заключения"]);
                        command.Parameters.AddWithValue("@Дата_начала_выполнения", row["Дата начала выполнения"]);
                        command.Parameters.AddWithValue("@Предмет_договора", row["Предмет договора"]);
                        command.Parameters.AddWithValue("@Сумма_договора", row["Сумма договора"]);
                        command.Parameters.AddWithValue("@Этап_договора", row["Этап договора"]);
                        command.Parameters.AddWithValue("@Дата_окончания_выполнения", row["Дата окончания выполнения"]);

                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
            }
        }

        private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
        {
            DataTable dataTable = (Table.ItemsSource as DataView)?.Table;

            if (dataTable != null)
            {
                SaveChangesToDatabase(dataTable);
                MessageBox.Show("Изменения успешно сохранены", "Сохранение изменений", MessageBoxButton.OK);
            }
            else
            {
                MessageBox.Show("Ошибка: не удалось получить данные для сохранения", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }
    }
}

