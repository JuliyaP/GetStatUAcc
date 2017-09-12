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
        // private static NpgsqlDataAdapter da = null;   
        // static string Test { get; set; } = "Test";
        static string Connstring { get; set; } = String.Format($"Server={appSet["Server"]};Port={appSet["Port"]};" +
                $"User Id={appSet["UserId"]}; Password= {appSet["Password"]};Database={appSet["Database"]};CommandTimeout=320;");
        // BackgroundWorker _worker;

        public MainWindow()
        {
            InitializeComponent();
            //  _worker = new BackgroundWorker();
            //  _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            //  _worker.WorkerReportsProgress = true;
            // _worker.DoWork += new DoWorkEventHandler(GetStatForAllUsers);

        }

        private void GetStatBtn_Click(object sender, RoutedEventArgs e)
        {
            Thread th = new Thread(GetStatForAllUsers);
            th.Start();
            //  _worker.RunWorkerAsync();
        }


        /// <summary>
        /// Получить статистику по всем пользователям ЛК
        /// </summary>
        public void GetStatForAllUsers()
        {

            Write("Начало обработки. Открытие соединения.");
            //   Thread.Sleep(TimeSpan.FromSeconds(5));
            cn.ConnectionString = Connstring;
            try
            {
                PathToSave.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)delegate () { PathToSave.Text = "Соед.открыто."; });
                cn.Open();
                WriteLine("  =>  Успешно");
                string sql = string.Format(@"SELECT  ""Login"" as ""Логин"",""INN"" as ""ИНН"", ""OGRN"" as ""ОГРН"" ,""Company"" as ""Название"",""MarketMembersTypes"" as ""Тип""FROM ""UserAccount"" where  ""IsActive"" = 'true'");
                NpgsqlCommand cmd = new NpgsqlCommand(sql, cn)
                {
                    CommandTimeout = 0
                };
                Write("Формирование выборки\n{0}\n..... Ждите\n", sql);
                //  NpgsqlDataReader dr = cmd.ExecuteReader();
                using (NpgsqlDataReader dr = cmd.ExecuteReader())
                {

                    if (dr != null && dr.HasRows)
                    {

                        WriteLine("Выборка сформирована. Записей => {0}", dr.RecordsAffected);
                        ExportToExcel(dr);
                    }
                    else
                    {
                        WriteLine("Выборка пуста");
                    }
                    WriteLine("Конец обработки");

                }
            }
            catch (Exception ex)
            {
                WriteLine(ex.Message);
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                {
                    cn.Close();
                }
            }
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
            int j = 0;
            List<UserAccount> customers = new List<UserAccount>();
            while (dr.Read())
            {
                try
                {
                    //UserAccount newItem = new UserAccount()
                    //{
                    //    Login = dr.GetString(0),
                    //    INN = dr.GetString(1),
                    //    OGRN = dr.GetString(2),
                    //    Company = dr.GetString(3),
                    //    MarketMembersTypes = dr.GetString(4)
                    //};


                    // customers.Add(newItem);


                    //foreach (var r in customers)
                    //{
                        i++;
                        ObjWorkSheet.Cells[i, 1] = dr.GetString(0);
                        ObjWorkSheet.Cells[i, 2] = dr.GetString(1);
                        ObjWorkSheet.Cells[i, 2].NumberFormat = "0";
                        ObjWorkSheet.Cells[i, 3] = dr.GetString(2);
                        ObjWorkSheet.Cells[i, 3].NumberFormat = "0";
                        ObjWorkSheet.Cells[i, 4] = dr.GetString(3);
                        ObjWorkSheet.Cells[i, 5] = dr.GetString(4);
                    //}


                }
                catch (Exception ex)
                {

                }

            }

            //foreach (var r in customers)
            //{
            //    i++;
            //    ObjWorkSheet.Cells[i, 1] = r.Login;
            //    ObjWorkSheet.Cells[i, 2] = r.INN;
            //    ObjWorkSheet.Cells[i, 2].NumberFormat = "0";
            //    ObjWorkSheet.Cells[i, 3] = r.OGRN;
            //    ObjWorkSheet.Cells[i, 3].NumberFormat = "0";
            //    ObjWorkSheet.Cells[i, 4] = r.MarketMembersTypes;
            //}

            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;

        }

    }
}
