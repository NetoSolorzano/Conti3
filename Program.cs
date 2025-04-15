using System;
using System.Data;
using System.Windows.Forms;

namespace Conti3
{
    internal static class Program
    {
        #region Variables globales
        public static string vg_user = ""; // codigo del usuario
        public static string vg_nuse = ""; // nombre del usuario
        public static string vg_cliente = ""; // nombre del usuario
        public static string vg_ipwan = "";    // direc wan del cliente

        public static DataTable dt_definic = new DataTable();    // definiciones
        public static DataTable dt_enlaces = new DataTable();    // enlaces
        #endregion

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new login());
        }
    }
}
