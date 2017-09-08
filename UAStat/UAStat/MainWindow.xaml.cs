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
using System.Windows.Data;
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

        //void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        //{
        //    PBar.Value = e.ProgressPercentage;
        //}
        /// <summary>
        /// Получить статистику по всем пользователям ЛК
        /// </summary>
        public void GetStatForAllUsers()
        {   
            
            Write("Начало обработки. Открытие соединения.");
            Thread.Sleep(TimeSpan.FromSeconds(5));
           // cn.ConnectionString = Connstring;
            try
            {
                PathToSave.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)delegate () { PathToSave.Text = "Соед.открыто."; });           
                cn.Open();
                WriteLine("  =>  Успешно");              
                string sql = string.Format(@"SELECT  ""Login"" as ""Логин"",""INN"" as ""ИНН"", ""OGRN"" as ""ОГРН"" ,""Company"" as ""Название"",""MarketMembersTypes"" as ""Тип""FROM ""UserAccount"" where  ""IsActive"" = 'true'");             
                NpgsqlCommand cmd = new NpgsqlCommand(sql, cn);
                cmd.CommandTimeout = 0;
                Write("Формирование выборки\n{0}\n..... Ждите\n", sql);
                NpgsqlDataReader dr = cmd.ExecuteReader();
                                         
                if (dr != null && dr.HasRows)
                {
                    WriteLine("Выборка сформирована. Записей => {0}", dr.RecordsAffected);         
                }
                else
                {
                    WriteLine("Выборка пуста");
                }
                WriteLine("Конец обработки");               
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

    }
}
