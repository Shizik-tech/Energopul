using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MySql.EntityFrameworkCore;

namespace Energopul
{
    public partial class AuthWindow : Window
    {
        public AuthWindow()
        {
            InitializeComponent();
        }
        private void AuthBtn_Click(object sender, RoutedEventArgs e)
        {
            string username = TxtUsername.Text;
            string password = TxtPassword.Password;

            if (username == "" && password == "")
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                Close();
            }
            else
            {
                MessageBox.Show("Неверные данные пользователя. Пожалуйста, повторите попытку.", "Ошибка аутентификации", MessageBoxButton.OK);
            }
        }
    }
}  
