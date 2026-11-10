using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PKNK_CNPM
{
    public static class AGI_DB
    {
        public static string PathDataBase = File.ReadAllText("FilePathDataBase.txt");
    }

    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new frmLoginPage());frmHomeScreen

            Application.Run(new frmHomeScreen());
        }
    }
}