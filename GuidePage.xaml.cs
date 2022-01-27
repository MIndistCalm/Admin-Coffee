using System;
using System.IO;
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
    /// Логика взаимодействия для GuidePage.xaml
    /// </summary>
    class GridItem
    {
        public int code { get; set; }
        public string title { get; set; }
        public string title2 { get; set; }
        public string title3 { get; set; }
        public string title4 { get; set; }
        public string title5 { get; set; }
        public string title6 { get; set; }
        public string title7 { get; set; }
        public string title8 { get; set; }
        public string title9 { get; set; }
        public string title10 { get; set; }
    }
public partial class GuidePage : Page
    {
        int guideList;
        List<GridItem> itemList = new List<GridItem>();
        public GuidePage()
        {
            InitializeComponent();
            UpdateClients();
            dataGrid_guides.ItemsSource = itemList;
        }
        public void UpdateClients()
        {
            guideList = 1;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Клиенты";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Фамилия";
            dataGrid_guides.Columns[2].Header = "Имя";
            dataGrid_guides.Columns[3].Header = "Отчество";
            dataGrid_guides.Columns[4].Header = "Дата рождения";
            dataGrid_guides.Columns[5].Header = "Постоянный клиент";
            dataGrid_guides.Columns[6].Header = "Телефон";
            dataGrid_guides.Columns[7].Header = "Почта";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Фамилия"], 
                        title2 = (string)dbReader["Имя"], 
                        title3 = (string)dbReader["Отчество"], 
                        title4 = (string)dbReader["Дата рождения"], 
                        title5 = (string)dbReader["Постоянный клиент"], 
                        title6 = (string)dbReader["Телефон"], 
                        title7 = (string)dbReader["Почта"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateSales()
        {
            guideList = 5;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Акции_и_скидки";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Наименование товара";
            dataGrid_guides.Columns[2].Header = "Старая цена";
            dataGrid_guides.Columns[3].Header = "Цена по скидке";
            dataGrid_guides.Columns[4].Header = "";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            Console.WriteLine(dataGrid_guides.Columns[1]);
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Наименование_товара"], 
                        title2 = (string)dbReader["Старая_цена"], 
                        title3 = (string)dbReader["Цена_по_скидке"]});
                }
            }
            
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdatePosts()
        {
            guideList = 2;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Должности";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Наименование должности";
            dataGrid_guides.Columns[2].Header = "Количество человек на текущей должности";
            dataGrid_guides.Columns[3].Header = "";
            dataGrid_guides.Columns[4].Header = "";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"],
                        title = (string)dbReader["Наименование должности"],
                        title2 = (string)dbReader["Количество человек на текущей должности"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateSmeta()
        {
            guideList = 3;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Смета";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Наименование товара";
            dataGrid_guides.Columns[2].Header = "";
            dataGrid_guides.Columns[3].Header = "";
            dataGrid_guides.Columns[4].Header = "";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Наименование_товара"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateMenu()
        {
            guideList = 4;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Меню";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Наименование продукта";
            dataGrid_guides.Columns[2].Header = "Цена";
            dataGrid_guides.Columns[3].Header = "Объём";
            dataGrid_guides.Columns[4].Header = "Состав";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";

            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Наименование_продукта"], 
                        title2 = (string)dbReader["Цена"], 
                        title3 = (string)dbReader["Объём"], 
                        title4 = (string)dbReader["Состав"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateFurniture_registr()
        {
            guideList = 6;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Реестр_мебели";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Наименование мебели";
            dataGrid_guides.Columns[2].Header = "Цвет мебели";
            dataGrid_guides.Columns[3].Header = "Производитель мебели";
            dataGrid_guides.Columns[4].Header = "";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";

            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Наименование_мебели"], 
                        title2 = (string)dbReader["Цвет_мебели"], 
                        title3 = (string)dbReader["Производитель_мебели"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateAdv()
        {
            guideList = 7;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Реклама";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Тип рекламы";
            dataGrid_guides.Columns[2].Header = "Цена";
            dataGrid_guides.Columns[3].Header = "";
            dataGrid_guides.Columns[4].Header = "";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() {
                        code = (int)dbReader["Код"],
                        title = (string)dbReader["Тип_рекламы"],
                        title2 = (string)dbReader["Цена"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateSuppliers()
        {
            guideList = 8;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Поставщики";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Наименование организации";
            dataGrid_guides.Columns[2].Header = "Наименование товаров";
            dataGrid_guides.Columns[3].Header = "Количество товара";
            dataGrid_guides.Columns[4].Header = "";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Наименование_организации"], 
                        title2 = (string)dbReader["Наименование_товаров"], 
                        title3 = (string)dbReader["Количество_товара"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateStaff()
        {
            guideList = 9;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Сотрудники";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Фамилия";
            dataGrid_guides.Columns[2].Header = "Имя";
            dataGrid_guides.Columns[3].Header = "Отчество";
            dataGrid_guides.Columns[4].Header = "Дата рождения";
            dataGrid_guides.Columns[5].Header = "Адрес проживания";
            dataGrid_guides.Columns[6].Header = "Телефон";
            dataGrid_guides.Columns[7].Header = "Почта";
            dataGrid_guides.Columns[8].Header = "Должность";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"], 
                        title = (string)dbReader["Фамилия"], 
                        title2 = (string)dbReader["Имя"], 
                        title3 = (string)dbReader["Отчество"], 
                        title4 = (string)dbReader["Дата"], 
                        title5 = (string)dbReader["Адрес"], 
                        title6 = (string)dbReader["Телефон"], 
                        title7 = (string)dbReader["Почта"], 
                        title8 = (string)dbReader["Должность"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateBranch()
        {
            guideList = 10;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Филиал";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Код Акции и скидки";
            dataGrid_guides.Columns[2].Header = "Код Клиенты";
            dataGrid_guides.Columns[3].Header = "Код Меню";
            dataGrid_guides.Columns[4].Header = "Код Поставщики";
            dataGrid_guides.Columns[5].Header = "Код Реестр мебели";
            dataGrid_guides.Columns[6].Header = "Код Реестр оборудования";
            dataGrid_guides.Columns[7].Header = "Код Реклама";
            dataGrid_guides.Columns[8].Header = "Код Сотрудники";
            dataGrid_guides.Columns[9].Header = "Код Товарооборот";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"],
                        title = dbReader["Код_Акции_и_скидки"].ToString(),
                        title2 = dbReader["Код_Клиенты"].ToString(),
                        title3 = dbReader["Код_Меню"].ToString(),
                        title4 = dbReader["Код_Поставщики"].ToString(),
                        title5 = dbReader["Код_Реестр_мебели"].ToString(),
                        title6 = dbReader["Код_Реестр_оборудования"].ToString(),
                        title7 = dbReader["Код_Реклама"].ToString(),
                        title8 = dbReader["Код_Сотрудники"].ToString(),
                        title9 = dbReader["Код_Товарооборот"].ToString()
                    });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }
        public void UpdateTrade_turnover()
        {
            guideList = 11;
            itemList.Clear();
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            string query = "SELECT * FROM Товарооборот";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();
            dataGrid_guides.Columns[1].Header = "Затраты";
            dataGrid_guides.Columns[2].Header = "Выручка";
            dataGrid_guides.Columns[3].Header = "Прибыль";
            dataGrid_guides.Columns[4].Header = "Наименование товара";
            dataGrid_guides.Columns[5].Header = "";
            dataGrid_guides.Columns[6].Header = "";
            dataGrid_guides.Columns[7].Header = "";
            dataGrid_guides.Columns[8].Header = "";
            dataGrid_guides.Columns[9].Header = "";
            if (dbReader.HasRows == true)
            {
                while (dbReader.Read())
                {
                    itemList.Add(new GridItem() { 
                        code = (int)dbReader["Код"],
                        title = (string)dbReader["Затраты"],
                        title2 = (string)dbReader["Выручка"],
                        title3 = (string)dbReader["Прибыль"],
                        title4 = (string)dbReader["Наименование товара"] });
                }
            }
            dataGrid_guides.Items.Refresh();
            dbReader.Close();
            dbConnection.Close();
        }

        private void clientsBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdateClients();
        }

        private void postsBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdatePosts();
        }

        private void smetaBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdateSmeta();
        }

        private void menuBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdateMenu();
        }
        private void salesBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdateSales();
        }
        private void suppliersBtn_Click(object sender, RoutedEventArgs e)
        {
            UpdateSuppliers();
        }
        private void furniture_registr_Btn_Click(object sender, RoutedEventArgs e)
        {
            UpdateFurniture_registr();
        }
        private void adv_Btn_Click(object sender, RoutedEventArgs e)
        {
            UpdateAdv();
        }
        private void staff_Btn_Click(object sender, RoutedEventArgs e)
        {
            UpdateStaff();
        }
        private void trade_turnover_Btn_Click(object sender, RoutedEventArgs e)
        {
            UpdateTrade_turnover();
        }
        private void branch_Btn_Click(object sender, RoutedEventArgs e)
        {
            UpdateBranch();
        }

        private void deleteBtn_Click(object sender, RoutedEventArgs e)
        {
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            GridItem item = (GridItem)dataGrid_guides.SelectedItem;
            if(item == null)
            {
                MessageBox.Show("Выберите поле!");
            }
            else
            {
                switch (guideList)
                {
                    case 1:
                        string query = "Delete FROM Клиенты WHERE Код = " + item.code + "";
                        OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateClients();
                        }
                        break;
                    case 2:
                        query = "Delete FROM Должности WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdatePosts();
                        }
                        break;
                    case 3:
                        query = "Delete FROM Смета WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateSmeta();
                        }
                        break;
                    case 4:
                        query = "Delete FROM Меню WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateMenu();
                        }
                        break;
                    case 5:
                        query = "Delete FROM Акции_и_скидки WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateSales();
                        }
                        break;
                    case 6:
                        query = "Delete FROM Реестр_мебели WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateFurniture_registr();
                        }
                        break;
                    case 7:
                        query = "Delete FROM Реклама WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateAdv();
                        }
                        break;
                    case 8:
                        query = "Delete FROM Поставщики WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateSuppliers();
                        }
                        break;
                    case 9:
                        query = "Delete FROM Сотрудники WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateStaff();
                        }
                        break;
                    case 10:
                        query = "Delete FROM Филиал WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateBranch();
                        }
                        break;
                    case 11:
                        query = "Delete FROM Товарооборот WHERE Код = " + item.code + "";
                        dbCommand = new OleDbCommand(query, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateTrade_turnover();
                        }
                        break;
                }
            }
            dbConnection.Close();
        }

        private int randomID()
        {
            DateTime centuryBegin = new DateTime(2021, 4, 29); //событие от которого рассчитывается количество тактов
            DateTime currentDate = DateTime.Now;
            return Math.Abs((int)currentDate.Ticks - (int)centuryBegin.Ticks);
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            GridItem item = (GridItem)dataGrid_guides.SelectedItem;
            int id = randomID();
            switch (guideList)
            {
                case 1:
                    string query = "INSERT INTO Клиенты (Код, [Фамилия], [Имя], [Отчество], [Дата рождения], [Постоянный клиент], [Телефон], [Почта]) VALUES (" + id + ", \" \", \" \", \" \", \" \", \" \", \" \", \" \")";
                    OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateClients();
                    }
                    break;
                case 2:
                    query = "INSERT INTO Должности (Код, [Наименование должности], [Количество человек на текущей должности]) VALUES (" + id + ", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdatePosts();
                    }
                    break;
                case 3:
                    query = "INSERT INTO Смета (Код, [Наименование_товара]) VALUES (" + id + ", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateSmeta();
                    }
                    break;
                case 4:
                    query = "INSERT INTO Меню (Код, [Наименование_продукта], [Цена], [Объём], [Состав]) VALUES (" + id + ", \" \", \" \", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateMenu();
                    }
                    break;
                case 5:
                    query = "INSERT INTO Акции_и_скидки (Код, [Наименование_товара], [Старая_цена], [Цена_по_скидке]) VALUES (" + id + ", \" \", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateSales();
                    }
                    break;
                case 6:
                    query = "INSERT INTO Реестр_мебели (Код, [Наименование_мебели], [Цвет_мебели], [Производитель_мебели]) VALUES (" + id + ", \" \", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateFurniture_registr();
                    }
                    break;
                case 7:
                    query = "INSERT INTO Реклама (Код, [Тип_рекламы], [Цена]) VALUES (" + id + ", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateAdv();
                    }
                    break;
                case 8:
                    query = "INSERT INTO Поставщики (Код, [Наименование_организации], [Наименование_товаров], [Количество_товара]) VALUES (" + id + ", \" \", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateSuppliers();
                    }
                    break;
                case 9:
                    query = "INSERT INTO Сотрудники (Код, [Фамилия], [Имя], [Отчество], [Дата], [Адрес], [Телефон], [Почта], [Должность]) VALUES (" + id + ", \" \", \" \", \" \", \" \", \" \", \" \", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateStaff();
                    }
                    break;
                case 10:
                    query = "INSERT INTO Филиал (Код, [Код_Акции_и_скидки], [Код_Клиенты], [Код_Меню], [Код_Поставщики], [Код_Реестр_мебели], [Код_Реестр_оборудования], [Код_Реклама], [Код_Сотрудники], [Код_Товарооборот]) VALUES (" + id + ", " + 0 + ", " + 0 + ", " + 0 + ", " + 0 + ", " + 0 + ", " + 0 + ", " + 0 + ", " + 0 + ", " + 0 + ")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка удаления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateBranch();
                    }
                    break;
                case 11:
                        
                    query = "INSERT INTO Товарооборот (Код, [Затраты], [Выручка], [Прибыль], [Наименование товара]) VALUES ("+ id +", \" \", \" \", \" \", \" \")";
                    dbCommand = new OleDbCommand(query, dbConnection);
                    if (dbCommand.ExecuteNonQuery() != 1)
                    {
                        MessageBox.Show("Ошибка добавления");
                    }
                    else
                    {
                        dbConnection.Close();
                        UpdateTrade_turnover();
                    }
                    break;
            }
            dbConnection.Close();
        }

        private void dataGrid_guides_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void airlinesBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void saveBtn_Click(object sender, RoutedEventArgs e)
        {
            string connectionString2 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=CafeDataBase.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString2);
            dbConnection.Open();
            GridItem item = (GridItem)dataGrid_guides.SelectedItem;
            int id = randomID();
            GridItem itemTemp;
            if (item == null)
            {
                MessageBox.Show("Выберите запись для сохранения!");
            }
            else
            {
                switch (guideList)
                {
                    case 1:
                        string queryDel = "Delete FROM Клиенты WHERE Код = " + item.code + "";
                        itemTemp = item;
                        OleDbCommand dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        string query = "INSERT INTO Клиенты (Код, [Фамилия], [Имя], [Отчество], [Дата рождения], [Постоянный клиент], [Телефон], [Почта]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "', '" + item.title4 + "', '" + item.title5 + "', '" + item.title6 + "', '" + item.title7 + "')";
                        OleDbCommand dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateClients();
                        }
                        break;
                    case 2:
                        queryDel = "Delete FROM Должности WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Должности (Код, [Наименование должности], [Количество человек на текущей должности]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdatePosts();
                        }
                        break;
                    case 3:
                        queryDel = "Delete FROM Смета WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Смета (Код,[Наименование_товара]) VALUES (" + itemTemp.code + ", '" + item.title + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateSmeta();
                        }
                        break;

                    case 4:
                        queryDel = "Delete FROM Меню WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Меню (Код, [Наименование_продукта], [Цена], [Объём], [Состав]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "', '" + item.title4 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateMenu();
                        }
                        break;

                    case 5:
                        queryDel = "Delete FROM Акции_и_скидки WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Акции_и_скидки (Код, [Наименование_товара], [Старая_цена], [Цена_по_скидке]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateSales();
                        }
                        break;

                    case 6:
                        queryDel = "Delete FROM Реестр_мебели WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Реестр_мебели (Код, [Наименование_мебели], [Цвет_мебели], [Производитель_мебели]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateFurniture_registr();
                        }
                        break;

                    case 7:
                        queryDel = "Delete FROM Реклама WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Реклама (Код, [Тип_рекламы], [Цена]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateAdv();
                        }
                        break;

                    case 8:
                        queryDel = "Delete FROM Поставщики WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Поставщики (Код, [Наименование_организации], [Наименование_товаров], [Количество_товара]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateSuppliers();
                        }
                        break;

                    case 9:
                        queryDel = "Delete FROM Сотрудники WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Сотрудники (Код, [Фамилия], [Имя], [Отчество], [Дата], [Адрес], [Телефон], [Почта], [Должность]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "', '" + item.title4 + "', '" + item.title5 + "', '" + item.title6 + "', '" + item.title7 + "', '" + item.title8 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateStaff();
                        }
                        break;

                    case 10:
                        queryDel = "Delete FROM Филиал WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Филиал (Код, [Код_Акции_и_скидки], [Код_Клиенты], [Код_Меню], [Код_Поставщики], [Код_Реестр_мебели], [Код_Реестр_оборудования], [Код_Реклама], [Код_Сотрудники], [Код_Товарооборот]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "', '" + item.title4 + "', '" + item.title5 + "', '" + item.title6 + "', '" + item.title7 + "', '" + item.title8 + "', '" + item.title9 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateBranch();
                        }
                        break;

                    case 11:
                        queryDel = "Delete FROM Товарооборот WHERE Код = " + item.code + "";
                        itemTemp = item;
                        dbCommand = new OleDbCommand(queryDel, dbConnection);
                        if (dbCommand.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        query = "INSERT INTO Товарооборот (Код, [Затраты], [Выручка], [Прибыль], [Наименование товара]) VALUES (" + itemTemp.code + ", '" + item.title + "', '" + item.title2 + "', '" + item.title3 + "', '" + item.title4 + "')";
                        dbCommand1 = new OleDbCommand(query, dbConnection);
                        if (dbCommand1.ExecuteNonQuery() != 1)
                        {
                            MessageBox.Show("Ошибка удаления");
                        }
                        else
                        {
                            dbConnection.Close();
                            UpdateTrade_turnover();
                        }
                        break;

                }
            }
            dbConnection.Close();
        }

        private void out_file(object sender, RoutedEventArgs e)
        {
            dataGrid_guides.SelectAllCells();
            dataGrid_guides.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, dataGrid_guides);
            dataGrid_guides.UnselectAllCells();
            var result = (string)Clipboard.GetData(DataFormats.Text);
            dynamic wordApp = null;
            try
            {
                var sw = new StreamWriter("export.doc");
                sw.WriteLine(result);
                sw.Close();
                //var proc = Process.Start("export.doc");
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                wordApp = Activator.CreateInstance(wordType);
                wordApp.Documents.Add(System.AppDomain.CurrentDomain.BaseDirectory + "export.doc");
                wordApp.ActiveDocument.Range.ConvertToTable(1, dataGrid_guides.Items.Count, dataGrid_guides.Columns.Count);
                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                }
                // ignored
            }
        }
    }
}
