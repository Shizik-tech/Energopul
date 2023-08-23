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
            throw new NotImplementedException();
        }
    }
}
