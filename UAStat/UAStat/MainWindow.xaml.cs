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
            using (UAStatContext db = new UAStatContext())
            {
                try
                {
                    List<UserAccount> us = db.Users.ToList();
                    ExportToExcel(us);
                }

                catch (Exception ex)
                {
                    OutputInfo($"Ошибка получения статистики. Подробнее: {ex.Message}");
                }
                finally
                {

                }
            }
        }


        public void OutputInfo(string message)
        {
            PathToSave.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)delegate () { PathToSave.Text = message; });
            WriteLine(message);
        }


        public void ExportToExcel(List<UserAccount> us)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            int i = 0;
  
                try
                {

                    i = InsertValuesToCells(us, ObjWorkSheet, i);
                }
                catch (Exception ex)
                {
                    OutputInfo($"Ошибка формирования excel файла. Подробнее: {ex.Message}");                   
                }
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;
        }

        private  int InsertValuesToCells(List<UserAccount> us, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, int i)
        {
            foreach (var user in us)
            {
                OutputInfo($"Запись в ячеку строка: {i}");
                i++;
                ObjWorkSheet.Cells[i, 1] = user.Login;
                ObjWorkSheet.Cells[i, 2] = user.INN;
                ObjWorkSheet.Cells[i, 2].NumberFormat = "0";
                ObjWorkSheet.Cells[i, 3] = user.OGRN;
                ObjWorkSheet.Cells[i, 3].NumberFormat = "0";
                ObjWorkSheet.Cells[i, 4] = user.MarketMembersTypes; 
                ObjWorkSheet.Cells[i, 5] = user.Company; 
            }
            return i;
        }
    }
}
