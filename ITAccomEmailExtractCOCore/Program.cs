using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ITAccomEmailExtractC
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DialogResult result;
            
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.SetCompatibleTextRenderingDefault(false);
            var loginForm = new FrmLoginAtt();
            result = loginForm.ShowDialog();
            if (result == DialogResult.OK)
            {
                Application.Run(new FrmEmailExport());
            }
            else
            {
                Application.Exit();
            }

          
        }
    }
}
