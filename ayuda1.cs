using System;
using System.Data;
using System.Configuration;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class ayuda1 : Form1
    {
        publicoConf conf = new publicoConf();
        libreria lib = new libreria();
        public string para1 = "";
        public string para2 = "";
        public string para3 = "";
        public string para4 = "";
        // Se crea un DataTable que almacenará los datos desde donde se cargaran los datos al DataGridView
        DataTable dtDatos = new DataTable();
        // string de conexion
        static string serv = ConfigurationManager.AppSettings["serv"].ToString();
        static string port = ConfigurationManager.AppSettings["port"].ToString();
        static string usua = ConfigurationManager.AppSettings["user"].ToString();
        static string cont = ConfigurationManager.AppSettings["pass"].ToString();
        static string data = ConfigurationManager.AppSettings["data"].ToString();
        string DB_CONN_STR = "server=" + serv + ";uid=" + usua + ";pwd=" + cont + ";database=" + data + ";";

        public ayuda1(string param1,string param2,string param3,string param4)
        {
            para1 = param1;              // 
            para2 = param2;              //
            para3 = param3;              //
            para4 = param4;              // 
            InitializeComponent();
        }
        private void ayuda1_Load(object sender, EventArgs e)
        {
            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml("#667d97");   //  "#656d77",   #f5510f", "#e76433"
            Bt_graba.Image = null;
            //this.BackColor = Color.FromArgb(conf.fondoPrinRojoE, conf.fondoPrinVerdeE, conf.fondoPriAzulE); // conf.fondoPrinBrilloE, 
            this.Text = "PROVEEDOR NUEVO";
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            //
            Tx_codigo.MaxLength = 6;
            Tx_codigo.CharacterCasing = CharacterCasing.Upper;
            Tx_codigo.ReadOnly = true;
            //Tx_codigo.BackColor = ColorTranslator.FromHtml("#d5f2de");
            Tx_nombre.MaxLength = 50;
            Tx_nombre.CharacterCasing = CharacterCasing.Upper;
            //Tx_nombre.BackColor = ColorTranslator.FromHtml("#d5f2de");
            tx_ruc.MaxLength = 11;
            tx_cuenta.MaxLength = 20;
            //
            ReturnValueA = new string[4] { "", "", "", "" };
        }
        private void ayuda1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        public string ReturnValue1 { get; set; }
        public string ReturnValue0 { get; set; }
        public string ReturnValue2 { get; set; }
        public string[] ReturnValueA { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Tx_nombre.Text != "")
            {
                var aa = MessageBox.Show("Confirma que desea grabar?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aa == DialogResult.Yes)
                {
                    graba();
                    ReturnValue0 = Tx_codigo.Text;
                    ReturnValue1 = Tx_nombre.Text;
                    ReturnValue0 = Tx_codigo.Text;
                    //
                    ReturnValueA[0] = Tx_codigo.Text;
                    ReturnValueA[1] = Tx_nombre.Text;
                    ReturnValueA[2] = tx_ruc.Text;      // ruc
                    ReturnValueA[3] = tx_cuenta.Text;   // cuenta
                }
            }
            this.Close();
        }

        private void Tx_codigo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (Tx_codigo.Text != "")
            {
                bool existe;
                using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                {
                    try
                    {
                        conn.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message,"Error de conexión",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        Application.Exit();
                    }
                    using (MySqlCommand micon = new MySqlCommand("select * from anag_for where idanagrafica=@codi", conn))
                    {
                        micon.Parameters.AddWithValue("@codi", Tx_codigo.Text.Trim());
                        using (MySqlDataReader dr = micon.ExecuteReader())
                        {
                            if (dr.HasRows == true) existe = true;
                            else existe = false;
                        }
                    }
                }
                if (existe == true)
                {
                    MessageBox.Show("Ya existe el código", "Error en código", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Tx_codigo.Clear();
                    Tx_nombre.Clear();
                    Tx_nombre.Focus();
                }
            }
        }
        private void graba()
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                string metela = "insert into anagrafiche " +
                    "(RagioneSociale,IDCategoria,stato,CodiceFiscale,ContoCorrente," +
                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) " +
                    "values (@nomb,'FOR',1,@ruc,@cta," +
                    "@vap,@asd,now(),@ipl,@ipw,@nbna)";
                using (MySqlCommand micon = new MySqlCommand(metela, conn))
                {
                    micon.Parameters.AddWithValue("@nomb", Tx_nombre.Text.Trim());
                    micon.Parameters.AddWithValue("@ruc", tx_ruc.Text);
                    micon.Parameters.AddWithValue("@cta", tx_cuenta.Text);
                    micon.Parameters.AddWithValue("@vap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                    micon.Parameters.AddWithValue("@asd", Program.vg_user);
                    micon.Parameters.AddWithValue("@ipl", lib.iplan());
                    micon.Parameters.AddWithValue("@ipw", Conti3.Program.vg_ipwan);
                    micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                    micon.ExecuteNonQuery();
                }
                //
                using (MySqlCommand micon = new MySqlCommand("select idanagrafica from anagrafiche where idcategoria='FOR' order by id desc limit 1", conn))
                {
                    using (MySqlDataReader dr = micon.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            Tx_codigo.Text = dr.GetString(0);
                        }
                    }
                }
            }
        }
    }
}
