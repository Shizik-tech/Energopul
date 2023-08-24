using System;
using System.Collections.Generic;
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
                string user = TxtUsername.Text;
                string pass = TxtPassword.Password;
                AuthWindow authWindow = new AuthWindow();
                authWindow.ReceivedUser = user;
                authWindow.ReceivedPass = pass;
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
