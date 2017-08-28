﻿using System;
using System.Collections.Generic;
using System.Data;
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
using Npgsql;

namespace UAStat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static NpgsqlConnection cn = new NpgsqlConnection();
       // private static NpgsqlDataAdapter da = null;   
        static string  test = "Test";
        static string  connstring = String.Format($"Server={test};Port={test};" +
                $"User Id={test}; Password= {test};Database={test};CommandTimeout=320;");

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GetStatBtn_Click(object sender, RoutedEventArgs e)
        {        
            GetStatForAllUsers();
        }

        /// <summary>
        /// Получить статистику по всем пользователям ЛК
        /// </summary>
        public void GetStatForAllUsers()
        {
            cn.ConnectionString = connstring;
            try
            {
                Console.Write("Начало обработки. Открытие соединения.");
                cn.Open();
                Console.WriteLine("  =>  Успешно");              
                string sql = string.Format(@"SELECT  ""Login"" as ""Логин"",""INN"" as ""ИНН"", ""OGRN"" as ""ОГРН"" ,""Company"" as ""Название"",""MarketMembersTypes"" as ""Тип""FROM ""UserAccount"" where  ""IsActive"" = 'true'");             
                NpgsqlCommand cmd = new NpgsqlCommand(sql, cn);
                cmd.CommandTimeout = 0;
                Console.Write("Формирование выборки\n{0}\n..... Ждите\n", sql);
                NpgsqlDataReader dr = cmd.ExecuteReader();
                                         
                if (dr != null && dr.HasRows)
                {
                    Console.WriteLine("Выборка сформирована. Записей => {0}", dr.RecordsAffected);         
                }
                else
                {                    
                    Console.WriteLine("Выборка пуста");
                }
                Console.WriteLine("Конец обработки");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                {
                    cn.Close();
                }
            }

        }
    }
}
