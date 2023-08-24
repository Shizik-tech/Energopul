using System;
using System.Collections.Generic;
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

namespace Energopul
{
    public partial class AuthDropWindow : Window
    {
        public AuthDropWindow()
        {
            InitializeComponent();
        }

        private void AuthDropBtn_Click(object sender, RoutedEventArgs e)
        {
            if (Spec.Text == "Энерго")
            {
                var authWindow = new AuthWindow();
                var dataToSave = TxtUsername.Text + Environment.NewLine + TxtPassword.Password;
                const string filePath = "data.txt"; // Укажите путь к файлу

                try
                {
                    File.WriteAllText(filePath, dataToSave);
                    MessageBox.Show("Успешно сброшены данные");
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Произошла ошибка при сохранении данных: " + ex.Message);
                }
                
                authWindow.Show();
                Close();
            }
            else
                MessageBox.Show("Введено неверное спецслово", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            var authWindow = new AuthWindow();
            authWindow.Show();
            Close();
        }
    }
}
