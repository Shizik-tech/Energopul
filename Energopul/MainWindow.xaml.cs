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
            InitializeComboBox();
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

        private void InitializeComboBox()
        {
            Period.Items.Add("Поквартально");
            Period.Items.Add("Полугодично");
            Period.Items.Add("Ежегодно");
            Period.Items.Add("Двухгодично");
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

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (fileInfo.Exists)
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
                MessageBox.Show("Экспорт данных успешно выполнен", "Экспорт", MessageBoxButton.OK);
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
                        *
                    FROM
                        Contracts 
                    WHERE
                        Дата_заключения BETWEEN @start_date AND @end_date_rout";

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
                        commandText = "INSERT INTO Contracts (Название, ИНН, Номер, Дата_заключения, Дата_окончания, Предмет_договора, Сумма, Этап, Сроки_по_этапам) VALUES (@Название, @ИНН, @Номер, @Дата_заключения, @Дата_окончания, @Предмет_договора, @Сумма, @Этап, @Сроки_по_этапам)";
                    }
                    else if (row.RowState == DataRowState.Modified)
                    {
                        // Операция UPDATE
                        commandText = "UPDATE Contracts SET Название = @Название, ИНН = @ИНН, Номер = @Номер, Дата_заключения = @Дата_заключения, Дата_окончания = @Дата_окончания, Предмет_договора = @Предмет_договора, Сумма = @Сумма, Этап = @Этап, Сроки_по_этапам = @Сроки_по_этапам WHERE id = @id";
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
                        command.Parameters.AddWithValue("@Номер", row["Номер"]);
                        command.Parameters.AddWithValue("@Дата_заключения", row["Дата_заключения"]);
                        command.Parameters.AddWithValue("@Дата_окончания", row["Дата_окончания"]);
                        command.Parameters.AddWithValue("@Предмет_договора", row["Предмет_договора"]);
                        command.Parameters.AddWithValue("@Сумма", row["Сумма"]);
                        command.Parameters.AddWithValue("@Этап", row["Этап"]);
                        command.Parameters.AddWithValue("@Сроки_по_этапам", row["Сроки_по_этапам"]);

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
            ButtonShow_Click(sender, e);
        }
    }
}

