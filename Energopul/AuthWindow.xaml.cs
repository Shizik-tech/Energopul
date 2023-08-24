using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
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
        public string ReceivedUser { get; set; }
        public string ReceivedPass { get; set; }

        public AuthWindow()
        {
            InitializeComponent();
        }


        private void AuthBtn_Click(object sender, RoutedEventArgs e)
        {
            const string filePath = "data.txt"; // Укажите путь к файлу

            try
            {
                if (File.Exists(filePath))
                {
                    var lines = File.ReadAllLines(filePath);

                    if (lines.Length >= 2)
                    {
                        var user = lines[0];
                        var password = lines[1];
                        if (TxtUsername.Text == user && TxtPassword.Password == password)
                        {
                            var mainWindow = new MainWindow();
                            mainWindow.Show();
                            Close();
                        }
                        else
                            MessageBox.Show("Неверный логин или пароль", "Ошибка", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        MessageBox.Show("Нет данных о логине или пароле", "Ошибка", 
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Нет файла с данными", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("Произошла ошибка при загрузке файла: " + ex.Message, "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DropBtn_Click(object sender, RoutedEventArgs e)
        {
            var authDropWindow = new AuthDropWindow();
            authDropWindow.Show();
            Close();
        }
    }
}