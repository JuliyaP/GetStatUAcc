using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using static System.Diagnostics.Debug;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Data.SqlClient;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Npgsql;
using System.Threading;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Configuration;
using AutoMapper;

namespace UAStat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static NpgsqlConnection cn = new NpgsqlConnection();
        static NameValueCollection appSet = ConfigurationSettings.AppSettings;
        static string Connstring { get; set; } = String.Format($"Server={appSet["Server"]};Port={appSet["Port"]};" +
                $"User Id={appSet["UserId"]}; Password= {appSet["Password"]};Database={appSet["Database"]};CommandTimeout=320;");

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GetStatBtn_Click(object sender, RoutedEventArgs e)
        {
            Thread th = new Thread(GetStatForAllUsers);
            th.Start();
        }


        /// <summary>
        /// Получить статистику по всем пользователям ЛК
        /// </summary>
        public void GetStatForAllUsers()
        {
            Write("Начало обработки. Открытие соединения.");
            cn.ConnectionString = Connstring;
            try
            {
                OutputInfo("Открытие соединения..");
                cn.Open();
                OutputInfo("Открытие соединения..Успешно.");
                string sql = string.Format(@"SELECT  ""Login"" as ""Логин"",""INN"" as ""ИНН"", ""OGRN"" as ""ОГРН"" ,""Company"" as ""Название"",""MarketMembersTypes"" as ""Тип""FROM ""UserAccount"" where  ""IsActive"" = 'true'");
                NpgsqlCommand cmd = new NpgsqlCommand(sql, cn)
                {
                    CommandTimeout = 0
                };
                OutputInfo("Формирование выборки.");
                using (NpgsqlDataReader dr = cmd.ExecuteReader())
                {

                    if (dr != null && dr.HasRows)
                    {
                        OutputInfo("Выборка сформирована.");
                        ExportToExcel(dr);
                    }
                    else
                    {
                        OutputInfo("Выборка пуста");
                    }
                    OutputInfo("Конец обработки");

                }
            }
            catch (Exception ex)
            {
                OutputInfo($"Ошибка получения статистики. Подробнее: {ex.Message}");
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                {
                    cn.Close();
                }
            }
        }


        public void OutputInfo(string message)
        {
            PathToSave.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)delegate () { PathToSave.Text = message; });
            WriteLine(message);
        }


        public void ExportToExcel(NpgsqlDataReader dr)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            int i = 0;
            while (dr.Read())
            {
                try
                {

                    i = InsertValuesToCells(dr, ObjWorkSheet, i);
                }
                catch (Exception ex)
                {
                    OutputInfo($"Ошибка формирования excel файла. Подробнее: {ex.Message}");                   
                }
            }
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;
        }

        private  int InsertValuesToCells(NpgsqlDataReader dr, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, int i)
        {
            OutputInfo($"Запись в ячеку строка: {i}");
            i++;
            ObjWorkSheet.Cells[i, 1] = dr.GetString(0);
            ObjWorkSheet.Cells[i, 2] = dr.GetString(1);
            ObjWorkSheet.Cells[i, 2].NumberFormat = "0";
            ObjWorkSheet.Cells[i, 3] = dr.GetString(2);
            ObjWorkSheet.Cells[i, 3].NumberFormat = "0";
            ObjWorkSheet.Cells[i, 4] = dr.GetString(3);
            ObjWorkSheet.Cells[i, 5] = dr.GetString(4);
            return i;
        }
    }
}
