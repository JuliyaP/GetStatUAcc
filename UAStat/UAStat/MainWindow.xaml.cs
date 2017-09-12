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
                    //List<MyClass> results = new List<MyClass>();

                    //while (dr.Read())
                    //{
                    //    MyClass newItem = new MyClass();

                    //    newItem.Id = dr.GetInt32(0);
                    //    newItem.TypeId = dr.GetInt32(1);
                    //    newItem.AllowedSMS = dr.GetBoolean(2);
                    //    newItem.TimeSpan = dr.GetString(3);
                    //    newItem.Price = dr.GetDecimal(4);
                    //    newItem.TypeName = dr.GetString(5);

                    //    results.Add(newItem);
                    //}
                }

         //       List<UserAccount> customers = new List<UserAccount>();// dr.AutoMap<UserAccount>()
                                                    // .ToList();
          



             
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
                    UserAccount newItem = new UserAccount();

                        newItem.Login = dr.GetString(0);
                       //newItem.INN = dr.GetInt32(1);
                    newItem.OGRN = dr.GetString(2);
                 //   newItem.OGRN = dr.GetString(3);
                    newItem.Company = dr.GetString(4);
                    newItem.MarketMembersTypes = dr.GetString(5);

                    customers.Add(newItem);
                }
                catch { }

            }
            //do
            //{

            // result += "<h1>Таблица №" + i + "</h1>";
            //DataTable schemaTable = dr.GetSchemaTable();



            //    foreach (DataRow row in schemaTable.Rows)
            //{
            //    i++;
            //    foreach (DataColumn column in schemaTable.Columns)
            //    {
            //        j++;
            //        ObjWorkSheet.Cells[i, j + 1] = column.ColumnName.ToString();
            //        //Console.WriteLine(String.Format("{0} = {1}",
            //        //   column.ColumnName, row[column]));
            //    }
            //}





            //while (dr.Read())
            //    {
            //       // result += "<li>";
            //        // Получить все поля строки
            //        for (int field = 0; field < dr.FieldCount; field++)
            //        {
            //            ObjWorkSheet.Cells[i, field + 1] = dr.GetName(field).ToString(); ;
            //            //result += "<b>" + reader.GetName(field).ToString() + "</b>" + ": " +
            //              //  reader.GetValue(field).ToString() + "<br>";
            //        }
            //      //  result += "</li>";
            //    }
            ////}
            //while (dr.NextResult());

            //while (dr.Read())
            //{
            //    int i = 1;
            //    for (int field = 0; field < dr.FieldCount; field++)
            //    {
            //        ObjWorkSheet.Cells[i, field+1] = dr.GetName(field).ToString(); ;
            //         //   dr.GetName(field).ToString();
            //        //dr += "<b>" + dr.GetName(field).ToString() + "</b>" + ": " +
            //        //    dr.GetValue(field).ToString() + "<br>";
            //    }
            //    i++;
            //    //Значения [y - строка,x - столбец]
            //    //ObjWorkSheet.Cells[3, 1] = "11";
            //    //ObjWorkSheet.Cells[3, 2] = "122";
            //    //ObjWorkSheet.Cells[3, 3] = "333";
            //}
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;

        }

    }
}
