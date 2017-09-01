using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using static System.Diagnostics.Debug;

namespace UAStat
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
           this.DispatcherUnhandledException += new DispatcherUnhandledExceptionEventHandler(App_DispatcherUnchandledException);
        }
        void App_DispatcherUnchandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            WriteLine("____________DispatcherUnchandledException");
            e.Handled = true; // необработанное искл., как обработанное
        }
    }
}
