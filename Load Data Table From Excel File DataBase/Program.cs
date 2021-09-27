using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Load_Data_Table_From_Excel_File_DataBase
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
            Application.Run(new Form1());
        }
    }
}
