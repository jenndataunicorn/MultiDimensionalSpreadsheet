using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Linq;
using System.Windows.Forms;
using Ssepan.Utility;

namespace DataGridViewVirtualModeTest
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static Int32 Main(String[] args)
        {
            //default to fail code
            Int32 returnValue = -1;

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1(args));

                returnValue = 0;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
            return returnValue;
        }
    }
}
