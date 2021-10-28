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
using System.Data.OleDb;

namespace Admin_Coffee
{
    /// <summary>
    /// Логика взаимодействия для AdminLoginPage.xaml
    /// </summary>
    public partial class AdminLoginPage : Page
    {
        public AdminLoginPage()
        {
            InitializeComponent();
        }

        private void loginBtn_Click(object sender, RoutedEventArgs e)
        {
            if(login.Text != "" && password.Password != "")
            {
                string connectionString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
                OleDbConnection dbConnection = new OleDbConnection(connectionString);
                dbConnection.Open();
                string query = "SELECT * FROM Администраторы WHERE Логин = '" + login.Text + "' AND Пароль = '" + password.Password + "';";
                OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
                OleDbDataReader dbReader = dbCommand.ExecuteReader();
                if (dbReader.HasRows == true)
                {
                    NavigationService.Navigate(new GuidePage());
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль!");
                }
            }
            else
            {
                MessageBox.Show("Введите логин и пароль!");
            }
        }
    }
}
