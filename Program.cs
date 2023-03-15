using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 簡易倉儲系統
{
    internal static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main(string[] parameter)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (parameter.Length <= 0)
                Application.Run(new UserView());
            else if (parameter[0] == "BOSS")
                Application.Run(new ManagerView());
        }
    }
}
