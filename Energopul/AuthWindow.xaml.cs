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
            string filePath = "data.txt"; // Укажите путь к файлу

            try
            {
                if (File.Exists(filePath))
                {
                    string[] lines = File.ReadAllLines(filePath);

                    if (lines.Length >= 2)
                    {
                        string User = lines[0];
                        string Password = lines[1];
                        if (TxtUsername.Text == User && TxtPassword.Password == Password)
                        {
                            MainWindow mainWindow = new MainWindow();
                            mainWindow.Show();
                            Close();
                        }
                        else
                            MessageBox.Show("Не верный логин или пароль");

                    }
                    else
                    {
                        MessageBox.Show("Нет данных о логине или пароле.");
                    }
                }
                else
                {
                    MessageBox.Show("Нет файла с данными.");
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("Ошибка при загруске файла: " + ex.Message);
            }

            
        }

        private void DropBtn_Click(object sender, RoutedEventArgs e)
        {
            AuthDropWindow authDropWindow = new AuthDropWindow();
            authDropWindow.Show();
            Close();
        }
    }
}  
