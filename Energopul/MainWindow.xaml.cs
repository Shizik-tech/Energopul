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
using System.Data;
using System.Data.Common;
using System.Windows;
using System.Windows.Controls;
using System.Data.SQLite;
using System;
using System.Data;

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
            const string connectionString = "Data Source=EnergiserDB.db;Version=3;";
            if (connectionString == null) throw new ArgumentNullException(nameof(connectionString));

            using var connection = new SQLiteConnection(connectionString);
            connection.Open();

            const string query = "SELECT * FROM Contracts";

            using var adapter = new SQLiteDataAdapter(query, connection);
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
