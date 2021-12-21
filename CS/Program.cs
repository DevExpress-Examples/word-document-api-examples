using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RichEditDocumentServerAPIExample
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
            WindowsFormsSettings.SetPerMonitorDpiAware();
            Application.Run(new Form1());
        }
    }
}
