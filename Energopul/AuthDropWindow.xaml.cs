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
    /// <summary>
    /// Логика взаимодействия для AuthDropWindow.xaml
    /// </summary>
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
                AuthWindow authWindow = new AuthWindow();
                string dataToSave = TxtUsername.Text + Environment.NewLine + TxtPassword.Password;
                string filePath = "data.txt"; // Укажите путь к файлу

                try
                {
                    File.WriteAllText(filePath, dataToSave);
                    MessageBox.Show("Успешно.");
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Error saving data: " + ex.Message);
                }
                
                authWindow.Show();
                Close();
            }
            else
                MessageBox.Show("Не верное спец слово");
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            AuthWindow authWindow = new AuthWindow();
            authWindow.Show();
            Close();
        }
    }
}
