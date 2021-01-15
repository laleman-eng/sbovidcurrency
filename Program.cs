using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;

namespace SBO_VID_Currency
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US", false);
            //Thread.CurrentThread.CurrentUICulture
            //Thread.CurrentCulture.NumberFormat.NumberDecimalSeparator = ".";
            Console.WriteLine("CurrentCulture is {0}.", CultureInfo.CurrentCulture.Name);
            Console.WriteLine("CurrentUICulture is {0}.", CultureInfo.CurrentUICulture.Name);


            Application.Run(new ServiceForm());
        }
    }
}
