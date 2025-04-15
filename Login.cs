using System;
using System.Windows.Forms;
using System.Data;
using MySql.Data.MySqlClient;
using System.Configuration;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;

namespace Conti3
{
    public partial class login : Form
    {
        publicoConf conf = new publicoConf();
        libreria lib = new libreria();
        // conexion a la base de datos
        public static string serv = ConfigurationManager.AppSettings["serv"].ToString();// Decrypt(ConfigurationManager.AppSettings["serv"].ToString(), true);     // "solorsoft.com";
        public static string port = ConfigurationManager.AppSettings["port"].ToString();
        public static string usua = ConfigurationManager.AppSettings["user"].ToString();                    // "solorsof_rei";
        public static string cont = ConfigurationManager.AppSettings["pass"].ToString(); // Decrypt(ConfigurationManager.AppSettings["pass"].ToString(), true);     // "190969Sorol";
        public static string data = ConfigurationManager.AppSettings["data"].ToString();
        public static string ctl = ConfigurationManager.AppSettings["ConnectionLifeTime"].ToString();
        string DB_CONN_STR = "server=" + serv + ";uid=" + usua + ";pwd=" + cont + ";database=" + data + ";";
        public DataTable dt_enlaces = new DataTable();

        public login()
        {
            InitializeComponent();
        }
        private void login_Load(object sender, EventArgs e)
        {
            string tituloF = "CONTICASSA 3.0";
            lb_version.Text = "Versión " + System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
            lb_titulo.Text = "CENTRO DE SERVICIOS OMG" + Environment.NewLine + tituloF;
            lb_titulo.BackColor = System.Drawing.Color.White;
            Image salir = Properties.Resources.Close_32; // Image.FromFile("recursos/Close_32.png");
            //Image entrar = Image.FromFile("recursos/ok.png");
            //pictureBox1.Image = logo;
            Button2.Image = salir;
            Button2.ImageAlign = ContentAlignment.MiddleCenter;
            //Button1.Image = entrar;
            init();
            // jala datos de configuracion
            jaladatos();
            backgroundWorker1.RunWorkerAsync();     // 08/03/2023
            Tx_user.Focus();
        }
        private void init()
        {
            checkBox1.Visible = false;
            tx_newcon.Visible = false;
            tx_newcon.MaxLength = 10;
            //
            //this.BackColor = System.Drawing.ColorTranslator.FromHtml(Program.colbac);
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            // validamos los campos
            string usuari = Tx_user.Text.Trim();     // usuario
            if (usuari == "" || usuari == "USUARIO")
            {
                MessageBox.Show("Por favor, ingrese el nombre de usuario", "Atención");
                Tx_user.Focus();
                return;
            }
            if (Tx_pwd.Text.Trim() == "" || Tx_pwd.Text == "CLAVE")
            {
                MessageBox.Show("Por favor, ingrese la contraseña", "Atención");
                Tx_pwd.Focus();
                return;
            }
            if (Tx_user.Text != "USUARIO" && Tx_pwd.Text != "CLAVE")
            {
                string contra = lib.md5(Tx_pwd.Text);
                MySqlConnection cn = new MySqlConnection(DB_CONN_STR);
                if (lib.procConn(cn) == true)
                {
                    //validamos que el usuario y passw son los correctos
                    string query = "select a.bloqueado,a.local,a.nombre,a.tipuser,a.nivel " +
                        "from usuarios a " +
                        "where a.nom_user=@usuario and a.pwd_user=@contra";
                    MySqlCommand mycomand = new MySqlCommand(query, cn);
                    mycomand.Parameters.AddWithValue("@usuario", Tx_user.Text);
                    mycomand.Parameters.AddWithValue("@contra", contra);
                    MySqlDataReader dr = mycomand.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            if (dr.GetInt16(0) == 0)    //    .GetString(0) == "0"
                            {
                                Conti3.Program.vg_user = Tx_user.Text;
                                Conti3.Program.vg_nuse = dr.GetString(2);
                                dr.Close();
                                // cambiamos la contraseña si fue hecha
                                cambiacont();
                                // jala la ip wan del cliente
                                try
                                {
                                    Conti3.Program.vg_ipwan = lib.ipwan();
                                }
                                catch
                                {
                                    Conti3.Program.vg_ipwan = "";
                                }
                                // nos vamos al form principal
                                Program.vg_user = this.Tx_user.Text;
                                //main Main = new main();
                                Main main = new Main();
                                main.Show();
                                this.Hide();
                            }
                            else
                            {
                                dr.Close();
                                MessageBox.Show("El usuario esta Bloqueado!");
                                return;
                            }
                        }
                        dr.Close();
                    }
                    else
                    {
                        dr.Close();
                        MessageBox.Show("Usuario y/o Contraseña erronea", " Atención ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    mycomand.Dispose();
                }
                cn.Close();
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            const string mensaje = "Deseas salir del sistema?";
            const string titulo = "Confirma por favor";
            var result = MessageBox.Show(mensaje, titulo,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            { Close(); }
        }
        private static string Decrypt(string cipherString, bool useHashing)
        {
            byte[] keyArray;
            //get the byte code of the string
            byte[] toEncryptArray = Convert.FromBase64String(cipherString);
            System.Configuration.AppSettingsReader settingsReader = new AppSettingsReader();
            //Get your key from config file to open the lock!
            //string key = (string)settingsReader.GetValue("pass",typeof(String));   // SecurityKey
            string key = "8312@Sorol";
            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //release any resource held by the MD5CryptoServiceProvider
                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }
            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;
            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
        private void Tx_user_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                Tx_pwd.Focus();
            }
        }
        private void Tx_user_Enter(object sender, EventArgs e)
        {
            if (Tx_user.Text == "USUARIO")
            {
                Tx_user.Text = "";
                Tx_user.ForeColor = System.Drawing.Color.Black;
            }
        }
        private void Tx_user_Leave(object sender, EventArgs e)
        {
            if (Tx_user.Text.Trim() == "")
            {
                Tx_user.Text = "USUARIO";
                Tx_user.ForeColor = System.Drawing.Color.Gray;
            }
        }
        private void Tx_pwd_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                Button1.PerformClick();
            }
        }
        private void Tx_pwd_TextChanged(object sender, EventArgs e)
        {
            if (panel1.Visible == true)
            {
                if (Tx_pwd.Text.Trim() != "" && Tx_pwd.Text.Trim() != "CLAVE")
                {
                    checkBox1.Visible = true;
                    checkBox1.Checked = false;
                    tx_newcon.Visible = false;
                }
                else
                {
                    checkBox1.Visible = false;
                    checkBox1.Checked = false;
                    tx_newcon.Visible = false;
                }
            }
        }
        private void Tx_pwd_Enter(object sender, EventArgs e)
        {
            if (Tx_pwd.Text == "CLAVE")
            {
                Tx_pwd.Text = "";
                Tx_pwd.ForeColor = System.Drawing.Color.Black;
                Tx_pwd.UseSystemPasswordChar = true;
            }
        }
        private void Tx_pwd_Leave(object sender, EventArgs e)
        {
            if (Tx_pwd.Text.Trim() == "")
            {
                Tx_pwd.Text = "CLAVE";
                Tx_pwd.ForeColor = System.Drawing.Color.Gray;
                Tx_pwd.UseSystemPasswordChar = false;
            }
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                tx_newcon.Visible = true;
                tx_newcon.Focus();
            }
            else
            {
                tx_newcon.Text = "";
                tx_newcon.Visible = false;
            }
        }
        private void cambiacont()
        {
            if (checkBox1.Checked == true && tx_newcon.Text != "")
            {
                MySqlConnection cn = new MySqlConnection(DB_CONN_STR);
                if (lib.procConn(cn) == true)
                {
                    string consulta = "update usuarios set pwd_user=@npa where nom_user=@nus";
                    MySqlCommand micon = new MySqlCommand(consulta, cn);
                    micon.Parameters.AddWithValue("@npa", lib.md5(tx_newcon.Text));
                    micon.Parameters.AddWithValue("@nus", Tx_user.Text);
                    try
                    {
                        micon.ExecuteNonQuery();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en actualización del password", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                        return;
                    }
                    micon.Dispose();
                }
                cn.Close();
            }
        }
        private void jaladatos()
        {
            MySqlConnection cn = new MySqlConnection(DB_CONN_STR);
            if (lib.procConn(cn) == true)
            {
                string consulta = "SELECT b.cliente,b.ruc,b.igv,b.direcc,b.ctadetra,b.detra " +
                    "from baseconf b";
                MySqlCommand micon = new MySqlCommand(consulta, cn);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        // 

                    }
                    dr.Close();
                    micon.Dispose();
                }
            }
            cn.Close();
        }
        private void login_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
        }
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked == true) tx_newcon.Visible = true;
            else tx_newcon.Visible = false;
        }
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                using (MySqlCommand mico = new MySqlCommand("select id,formulario,campo,descrip,valor,param from enlaces", conn))
                {
                    using (MySqlDataAdapter da = new MySqlDataAdapter(mico))
                    {
                        da.Fill(dt_enlaces);
                        foreach (DataRow row in dt_enlaces.Rows)
                        {
                            /*
                            configuracion.dt_enlacesRow nr = setC.dt_enlaces.Newdt_enlacesRow();
                            nr.id = int.Parse(row.ItemArray[0].ToString());
                            nr.formulario = row.ItemArray[1].ToString();
                            nr.campo = row.ItemArray[2].ToString();
                            nr.descrip = row.ItemArray[3].ToString();
                            nr.valor = row.ItemArray[4].ToString();
                            nr.param = row.ItemArray[5].ToString();
                            setC.dt_enlaces.Adddt_enlacesRow(nr);
                            */
                        }
                    }
                }
            }
        }
    }
}
