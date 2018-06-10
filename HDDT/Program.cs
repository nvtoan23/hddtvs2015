using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace HDDT
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DateTime today = DateTime.Now;
            DateTime exp = new DateTime(2018, 10, 01);
            if (today.CompareTo(exp) == 1)
            {
                //MessageBox.Show("Hết hạn dùng");
                //Application.Exit();
            }
            else
            {
            
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new frmMain());
            }
        }

    }
}