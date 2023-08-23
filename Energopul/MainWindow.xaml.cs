using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
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
            FileInfo fileInfo = new FileInfo(fileName);

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
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Таблица1");
    
            IRow headerRow = sheet.CreateRow(0);
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                headerRow.CreateCell(col).SetCellValue(dataTable.Columns[col].ColumnName);
            }
    
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                IRow dataRow = sheet.CreateRow(row + 1);
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    dataRow.CreateCell(col).SetCellValue(dataTable.Rows[row][col].ToString());
                }
            }
    
            using (FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
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
    }
}
