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
            const string relativePath = "energopul.db";
            var fullPath = Path.Combine(AppContext.BaseDirectory, relativePath);
            return $"Data Source={fullPath};Version=3;";
        }

        private DataTable GetDataFromSqLite(string query)
        {
            using var connection = new SQLiteConnection(GetConnectionString());
            connection.Open();
            using var command = new SQLiteCommand(query, connection);
            using var adapter = new SQLiteDataAdapter(command);
            var dataTable = new DataTable();
            adapter.Fill(dataTable);
            return dataTable;
        }

        private void InitializeComboBox()
        {
            Period.Items.Add("Поквартально");
            Period.Items.Add("Полугодично");
            Period.Items.Add("Ежегодно");
            Period.Items.Add("Двухгодично");
            Period.Items.Add("За всё время");
        }

        private void LoadDataToDataGrid()
        {
            const string query = "SELECT * FROM Contracts;";
            var dataTable = GetDataFromSqLite(query);
            Table.ItemsSource = dataTable.DefaultView;
        }



        private void ExportDataTableToExcel(DataTable dataTable, string fileName)
        {
            var fileInfo = new FileInfo(fileName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (fileInfo.Exists)
            {
                var result = MessageBox.Show("Документ с таким названием уже существует. Перезаписать?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

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


            using (var package = new ExcelPackage(new FileInfo(fileName)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                for (var col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                }

                for (var row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (var col = 0; col < dataTable.Columns.Count; col++)
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
            using var doc = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());


            var table = new Table();


            var tblProperties = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "100%" });

            table.AppendChild(tblProperties);


            var headerRow = new TableRow();

            foreach (DataColumn column in dataTable.Columns)
            {
                var headerCell = new TableCell(new Paragraph(new Run(new Text(column.ColumnName))));
                headerRow.Append(headerCell);
            }

            table.Append(headerRow);


            foreach (DataRow dataRow in dataTable.Rows)
            {
                var dataTableRow = new TableRow();

                foreach (var item in dataRow.ItemArray)
                {
                    var dataCell = new TableCell(new Paragraph(
                        new Run(new Text(item.ToString()))));
                    dataTableRow.Append(dataCell);
                }

                table.Append(dataTableRow);
            }

            body.Append(table);
            MessageBox.Show("Экспорт данных успешно выполнен", "Экспорт", MessageBoxButton.OK);
        }

        private void ExportDataButton_Click(object sender, RoutedEventArgs e)
        {
            const string query = "SELECT * FROM Contracts";
            var dataTable = GetDataFromSqLite(query);
            ExportDataTableToExcel(dataTable, "Contracts.xlsx");
        }
        private void ExportDataToWordButton_Click(object sender, RoutedEventArgs e)
        {
            const string query = "SELECT * FROM Contracts";
            var dataTable = GetDataFromSqLite(query);
            ExportDataTableToWord(dataTable, "Contracts.docx");
        }

        private void ButtonShow_Click(object sender, EventArgs e)
        {
            var periodIndex = Period.SelectedIndex + 1; // Индексы ComboBox начинаются с 0
            var conn = new SQLiteConnection(GetConnectionString());
            using var cmd = conn.CreateCommand();
            var endDateRout = DateTime.Now.ToString("yyyy-MM-dd");

            string startDate;
            switch (periodIndex)
            {
                case 1:
                    startDate = DateTime.Now.AddMonths(-3).ToString("yyyy-MM-dd");
                    break;
                case 2:
                    startDate = DateTime.Now.AddMonths(-6).ToString("yyyy-MM-dd");
                    break;
                case 3:
                    startDate = DateTime.Now.AddYears(-1).ToString("yyyy-MM-dd");
                    break;
                case 4:
                    startDate = DateTime.Now.AddYears(-2).ToString("yyyy-MM-dd");
                    break;
                case 5:
                    startDate = "0000-00-00";
                    break;
                default:
                    MessageBox.Show("Неверный период");
                    return;
            }

            cmd.CommandText = $@"
                    SELECT
                        *
                    FROM
                        Contracts 
                    WHERE
                        Дата_заключения BETWEEN @start_date AND @end_date_rout";

            cmd.Parameters.AddWithValue("@start_date", startDate);
            cmd.Parameters.AddWithValue("@end_date_rout", endDateRout);

            var dataTable = new DataTable();
            using (var adapter = new SQLiteDataAdapter(cmd))
            {
                adapter.Fill(dataTable);
            }

            // Привязка данных к WPF DataGrid (Table)
            Table.ItemsSource = dataTable.DefaultView;
        }


        private void SaveChangesToDatabase(DataTable dataTable)
        {
            var connectionString = GetConnectionString();

            using var connection = new SQLiteConnection(connectionString);
            connection.Open();

            using var transaction = connection.BeginTransaction();
            using var command = connection.CreateCommand();
            try
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    if (row.RowState == DataRowState.Deleted)
                    {
                        // Delete logic
                        var originalId = row["id", DataRowVersion.Original];
                        if (originalId == DBNull.Value) continue;
                        command.CommandText = "DELETE FROM Contracts WHERE id = @id";
                        command.Parameters.Clear();
                        command.Parameters.AddWithValue("@id", originalId);
                        command.ExecuteNonQuery();
                    }
                    else if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                    {
                        // Insert or Update logic
                        command.Parameters.Clear();
                        command.Parameters.AddWithValue("@id", row["id"]);
                        command.Parameters.AddWithValue("@Название", row["Название"]);
                        command.Parameters.AddWithValue("@ИНН", row["ИНН"]);
                        command.Parameters.AddWithValue("@Номер", row["Номер"]);
                        command.Parameters.AddWithValue("@Дата_заключения", row["Дата_заключения"]);
                        command.Parameters.AddWithValue("@Дата_окончания", row["Дата_окончания"]);
                        command.Parameters.AddWithValue("@Предмет_договора", row["Предмет_договора"]);
                        command.Parameters.AddWithValue("@Сумма", row["Сумма"]);
                        command.Parameters.AddWithValue("@Этап", row["Этап"]);
                        command.Parameters.AddWithValue("@Сроки_по_этапам", row["Сроки_по_этапам"]);

                        if (row.RowState == DataRowState.Added)
                        {
                            command.CommandText = "INSERT INTO Contracts  (Название, ИНН, Номер, " +
                                                  "Дата_заключения, Дата_окончания, Предмет_договора, " +
                                                  "Сумма, Этап, Сроки_по_этапам) VALUES " +
                                                  "(@Название, @ИНН, @Номер, @Дата_заключения, @Дата_окончания, " +
                                                  "@Предмет_договора, @Сумма, @Этап, @Сроки_по_этапам)";
                        }
                        else if (row.RowState == DataRowState.Modified)
                        {
                            command.CommandText = "UPDATE Contracts SET Название = @Название, ИНН = @ИНН, " +
                                                  "Номер = @Номер, Дата_заключения = @Дата_заключения, " +
                                                  "Дата_окончания = @Дата_окончания, " +
                                                  "Предмет_договора = @Предмет_договора, Сумма = @Сумма, " +
                                                  "Этап = @Этап, Сроки_по_этапам = @Сроки_по_этапам WHERE id = @id";
                        }

                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Возникла ошибка: " + ex.Message, "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
        {
            var dataTable = (Table.ItemsSource as DataView)?.Table;

            if (dataTable != null)
            {
                SaveChangesToDatabase(dataTable);
                MessageBox.Show("Изменения успешно сохранены", "Сохранение изменений", 
                    MessageBoxButton.OK);
            }
            else
            {
                MessageBox.Show("Не удалось получить данные для сохранения", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ButtonShow_Click(sender, e);
        }

        private void SearchBtn_Click(object sender, RoutedEventArgs e)
        {
            var searchValue = Search.Text; // Получаем значение поиска из текстового поля

            if (string.IsNullOrWhiteSpace(searchValue))
            {
                MessageBox.Show("Введите значение для поиска", "Поиск", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            const string query = "SELECT * FROM Contracts WHERE Номер LIKE @searchValue";

            var dataTable = new DataTable();
            using (var connection = new SQLiteConnection(GetConnectionString()))
            using (var command = new SQLiteCommand(query, connection))
            {
                command.Parameters.AddWithValue("@searchValue", "%" + searchValue + "%"); 
                // Используем подстановку для частичного соответствия

                connection.Open();
                using (var adapter = new SQLiteDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }
            }

            Table.ItemsSource = dataTable.DefaultView; // Обновляем DataGrid с результатами поиска
        }

        private void ResetBtn_OnClick(object sender, RoutedEventArgs e)
        {
            LoadDataToDataGrid();
            MessageBox.Show("Настройки фильтра сброшены");
        }
    }
}

