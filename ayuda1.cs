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
        Finan_Egres OFegres = new Finan_Egres();
        libreria lib = new libreria();
        ccolores OColores = new ccolores();
        public string para1 = "";
        public string para2 = "";
        public string para3 = "";
        public string para4 = "";
        // Se crea un DataTable que almacenará los datos desde donde se cargaran los datos al DataGridView
        //DataTable dtDatos = new DataTable();
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
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                OFegres.jalacolores(conn, OColores, "provee");
            }
            OFegres.colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);
            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml(OColores.Fondo_boton_graba);
            Bt_graba.Image = null;
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
            tx_telef.MaxLength = 15;
            tx_mail.MaxLength = 50;
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
            // validaciones
            if (Tx_nombre.Text.Trim() == "")
            {
                MessageBox.Show("Debe ingresar el nombre","Atención",MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Tx_nombre.Focus();
                return;
            }
            if (tx_cuenta.Text.Trim() == "")
            {
                MessageBox.Show("Debe ingresar la cuenta bancaria", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                tx_cuenta.Focus();
                return;
            }
            if (tx_telef.Text.Trim() == "" && tx_mail.Text.Trim() == "")
            {
                MessageBox.Show("Debe ingresar el teléfono", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                tx_telef.Focus();
                return;
            }
            if (tx_mail.Text.Trim() == "" && tx_telef.Text.Trim() == "")
            {
                MessageBox.Show("Debe ingresar el correo electrónico", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                tx_mail.Focus();
                return;
            }

            //if (Tx_nombre.Text != "")
            {
                string err1 = ""; string err2 = "";
                using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                {
                    conn.Open();
                    string uno = "select count(id) from anagrafiche where RagioneSociale=@nomb";
                    using (MySqlCommand micon = new MySqlCommand(uno, conn)) 
                    {
                        micon.Parameters.AddWithValue("@nomb", Tx_nombre.Text);
                        using (MySqlDataReader dr = micon.ExecuteReader())
                        {
                            if (dr.Read()) 
                            {
                                if (dr.GetInt32(0) > 0) err1 = "El nombre del proveedor ya existe";
                            }
                        }
                    }
                    string dos = "";
                    if (tx_ruc.Text != "")
                    {
                        dos = "select RagioneSociale from anagrafiche where CodiceFiscale=@nruc";
                        using (MySqlCommand micon = new MySqlCommand(dos, conn))
                        {
                            micon.Parameters.AddWithValue("@nruc", tx_ruc.Text);
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.HasRows)
                                {
                                    if (dr.Read())
                                    {
                                        if (dr.GetString(0).Trim() != "") err2 = "El RUC ingresado ya existe" + Environment.NewLine +
                                                "correspode a: " + dr.GetString(0).Trim();
                                    }
                                }
                            }
                        }
                    }

                    if (err1 != "")
                    {
                        MessageBox.Show(err1,"Error!");
                        return;
                    }
                    if (err2 != "")
                    {
                        MessageBox.Show(err2, "Error!");
                        return;
                    } 
                }

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
                    "(RagioneSociale,IDCategoria,stato,CodiceFiscale,ContoCorrente,NumeroTel1,EMail," +
                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) " +
                    "values (@nomb,'FOR',1,@ruc,@cta,@tel1,@mail," +
                    "@vap,@asd,now(),@ipl,@ipw,@nbna)";
                using (MySqlCommand micon = new MySqlCommand(metela, conn))
                {
                    micon.Parameters.AddWithValue("@nomb", Tx_nombre.Text.Trim());
                    micon.Parameters.AddWithValue("@ruc", tx_ruc.Text);
                    micon.Parameters.AddWithValue("@cta", tx_cuenta.Text);
                    micon.Parameters.AddWithValue("@tel1", tx_telef.Text);
                    micon.Parameters.AddWithValue("@mail", tx_mail.Text);
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
