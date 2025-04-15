using System;
//using System.Collections.Generic;
//using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class Main : Form1
    {
        publicoConf conf = new publicoConf();

        #region variables
        string img_log1 = "";   //@"C:\omg-peru\imAGENES\images6.jpg";
        string nomclie = "";
        string dirclie = "";
        #endregion

        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";port=" + login.port + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data +
            ";ConnectionLifeTime=" + login.ctl + ";";

        public bool isResizeMode { get; private set; }

        public Main()
        {
            InitializeComponent();
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            try
            {
                conn.Open();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message,"Error de conexión al servidor");
                Application.Exit();
                return;
            }
            jalainfo(conn);                                 // jalamos los parametros, variables, etc
            Main_Load();                                // carga configuración del form
            cuadre();                                   // acomodo de botones verticales
            coloracion();                               // colores de botones y otros
            conn.Close();
        }

        private void Main_Load()
        {
            this.BackColor = Color.FromArgb(conf.fondoPrinRojoE, conf.fondoPrinVerdeE, conf.fondoPriAzulE); // conf.fondoPrinBrilloE, 

            Image logo1 = Image.FromFile(img_log1);
            Image salir = Properties.Resources.shut_down40;              // Image.FromFile("shut_down40.png");
            Image finan = Properties.Resources.inbox_done40;             // Image.FromFile("inbox_done40.png");
            Image webser = Properties.Resources.wifi40;                  // Image.FromFile("wifi40.png");
            Image valid = Properties.Resources.fileboard_checklist40;    // Image.FromFile("fileboard_checklist40.png");
            Image maest = Properties.Resources.database_system40;        // Image.FromFile("database_system40.png");
            Image panel = Properties.Resources.settings40;               // Image.FromFile("settings40.png");
            pictureBox1.Image = logo1;
            bt_salir.Image = salir;
            bt_finan.Image = finan;
            bt_serWeb.Image = webser;
            bt_validac.Image = valid;
            bt_maestras.Image = maest;
            bt_pcontrol.Image = panel;
            //
            tx_user.Text = Program.vg_user;         // código de usuario
            tx_nuser.Text = Program.vg_nuse;        // nombre de usuario
            tx_empresa.Text = nomclie;              // nombre de la organización
            //
            //pn_phor.Controls.Add(pn_menu);
            //pn_menu.Width = pn_phor.Width;  // - pn_acciones.Width;
            menuStrip1.Visible = true;
            pn_menu.Controls.Add(menuStrip1);
            menuStrip1.Dock = DockStyle.Top;
        }
        private void jalainfo(MySqlConnection conn)
        {
            string consulta = "select * from baseconf limit 1";
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                using (MySqlDataReader dr = micon.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            nomclie = dr.GetString("Cliente");              // nombre comercial
                            //rucclie = dr.GetString("Ruc");
                            dirclie = dr.GetString("direcc").Trim();
                            //rasclie = dr.GetString("razonsocial");          // nombre 
                            //tasaigv = dr.GetString("igv");
                            //ubigeoe = dr.GetString("referen1");             // ubigeo
                            //distemi = dr.GetString("distrit").Trim();
                            //provemi = dr.GetString("provin").Trim();
                            //urbemis = dr.GetString("referen2").Trim();      // urbanizacion
                            //depaemi = dr.GetString("depart").Trim();        // departamento
                            //urlemis = dr.GetString("urlCliente").Trim();    // web
                        }
                    }
                }
            }
            consulta = "select IDTabella,IDCodice,Descrizione,DescrizioneRid,Numero,sede,placa " +
                    "from descrittive order by IDTabella,IDCodice"; // where numero=1 
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                {
                    da.Fill(Program.dt_definic);
                }
            }
            consulta = "select * from enlaces";
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                {
                    da.Fill(Program.dt_enlaces);
                }
                DataRow[] rila = Program.dt_enlaces.Select("formulario = 'Main' and campo='imagen' and param='logo'");
                img_log1 = rila[0]["valor"].ToString();
            }
        }
        private void cuadre()
        {
            ControlBox = true;
            MaximizeBox = true;
            MinimizeBox = true;
            FormBorderStyle = FormBorderStyle.Sizable;  // FormBorderStyle.FixedSingle
            //Text = Program.tituloF;
            Left = Screen.PrimaryScreen.Bounds.Left;
            Top = Screen.PrimaryScreen.Bounds.Top;
            //Width = Screen.PrimaryScreen.Bounds.Width;
            //Height = Screen.PrimaryScreen.Bounds.Height;
            //
            bt_finan.Top = pictureBox1.Top + pictureBox1.Height + 0;
            bt_serWeb.Top = bt_finan.Top + bt_finan.Height + 0;
            bt_validac.Top = bt_serWeb.Top + bt_serWeb.Height + 0;
            //bt_almacen.Top = bt_pedidos.Top + bt_pedidos.Height + 0;
            bt_maestras.Top = bt_validac.Top + bt_validac.Height + 0;
            bt_pcontrol.Top = bt_maestras.Top + bt_maestras.Height + 0;

        }
        private void coloracion()
        {
            pn_centro.BackColor = Color.FromArgb(conf.fondoPrinRojoE,conf.fondoPrinVerdeE, conf.fondoPriAzulE); // conf.fondoPrinBrilloE,
            pn_pver.BackColor = Color.White; // Color.FromArgb(242, 243, 219);  // RGB
            bt_finan.BackColor = Color.White;
            bt_salir.BackColor = Color.White;
            bt_serWeb.BackColor = Color.White;
            bt_validac.BackColor = Color.White;
            bt_maestras.BackColor = Color.White;
            bt_pcontrol.BackColor = Color.White;
            pn_user.BackColor = Color.White;
            pn_menu.BackColor = Color.White;
            //pn_acciones.BackColor = Color.White;
        }

        #region botones_click   // menus

        #region finanzas
        private void bt_finan_Click(object sender, EventArgs e)
        {
            pic_icon_menu.Image = Properties.Resources.inbox_done20; // Resource.inbox_done20;
            menuStrip1.Items.Clear();
            menuStrip1.Items.Add("Ingresos", Properties.Resources.plus_circle20, fin_ingresos_Click);           // img_F1
            menuStrip1.Items.Add("Egresos", Properties.Resources.minus_circle20, fin_egresos_Click);             // img_F2
            menuStrip1.Items.Add("Gastos Almacén", Properties.Resources.truck_round20, fin_camion_Click);      // img_F3
            menuStrip1.Items.Add("Reportes", Properties.Resources.file_arrow_down20, fin_reportes_Click);           // img_F5
            //
            menuStrip1.Visible = true;
        }
        private void fin_ingresos_Click(object sender, EventArgs e)
        {
            Finan_Ingres ffe0 = new Finan_Ingres();
            ffe0.TopLevel = false;
            ffe0.Parent = this;
            ffe0.Top = pn_centro.Top ;
            ffe0.Left = pn_pver.Left ;
            pn_centro.Controls.Add(ffe0);
            if (this.Width < ffe0.Width + pn_centro.Left) this.Width = pn_centro.Left + ffe0.Width + 20;
            ffe0.Show();
        }
        private void fin_egresos_Click(object sender, EventArgs e)
        {
            Finan_Egres ffe1 = new Finan_Egres();
            ffe1.TopLevel = false;
            ffe1.Parent = this;
            ffe1.Left = pn_pver.Left + pn_pver.Width + 1;
            pn_centro.Controls.Add(ffe1);
            if (this.Width < ffe1.Right + ffe1.Left) this.Width = ffe1.Right + ffe1.Left;
            ffe1.Show();
        }
        private void fin_camion_Click(object sender, EventArgs e)
        {
            Finan_camion ffe1 = new Finan_camion();
            ffe1.TopLevel = false;
            ffe1.Parent = this;
            ffe1.Left = pn_pver.Left + pn_pver.Width + 1;
            pn_centro.Controls.Add(ffe1);
            if (this.Width < ffe1.Right + ffe1.Left) this.Width = ffe1.Right + ffe1.Left;
            ffe1.Show();
        }
        private void fin_reportes_Click(object sender, EventArgs e)
        {
            Finan_reps1 ffe1 = new Finan_reps1();
            ffe1.TopLevel = false;
            ffe1.Parent = this;
            ffe1.Left = pn_pver.Left + pn_pver.Width + 1;
            pn_centro.Controls.Add(ffe1);
            if (this.Width < ffe1.Right + ffe1.Left) this.Width = ffe1.Right + ffe1.Left;
            ffe1.Show();
        }
        #endregion

        #region servicios web
        private void bt_serWeb_Click(object sender, EventArgs e)
        {
            pic_icon_menu.Image = Properties.Resources.wifi40;
            menuStrip1.Items.Clear();
            // 
            menuStrip1.Items.Add("Gestión Usuarios Web", Properties.Resources.cloud20x14, web_usuarios_Click);
            menuStrip1.Visible = true;
        }
        private void web_usuarios_Click(object sender, EventArgs e)
        {
            usuarweb ffe0 = new usuarweb();
            ffe0.TopLevel = false;
            ffe0.Parent = this;
            ffe0.Top = pn_centro.Top;
            ffe0.Left = pn_pver.Left;
            pn_centro.Controls.Add(ffe0);
            if (this.Width < ffe0.Width + pn_centro.Left) this.Width = pn_centro.Left + ffe0.Width + 20;
            ffe0.Show();
        }
        #endregion

        #region validaciones de ingreso
        private void bt_validac_Click(object sender, EventArgs e)
        {
            pic_icon_menu.Image = Properties.Resources.fileboard_checklist40;
            pic_icon_menu.Tag = "VALIDACION";
            menuStrip1.Items.Clear();
            // 
            menuStrip1.Items.Add("Registros para Validar", Properties.Resources.fileboard_checklist20, vali_ingresos_Click);
            menuStrip1.Visible = true;
        }
        private void vali_ingresos_Click(object sender, EventArgs e)
        {
            egborrador ffe0 = new egborrador();
            ffe0.TopLevel = false;
            ffe0.Parent = this;
            ffe0.Top = pn_centro.Top;
            ffe0.Left = pn_pver.Left;
            pn_centro.Controls.Add(ffe0);
            if (this.Width < ffe0.Width + pn_centro.Left) this.Width = pn_centro.Left + ffe0.Width + 20;
            ffe0.Show();
        }
        #endregion

        #region maestras
        private void bt_maestras_Click(object sender, EventArgs e)
        {
            pic_icon_menu.Image = Properties.Resources.database_system20;   // Resource1.database_system20;
            menuStrip1.Items.Clear();
            menuStrip1.Items.Add("Proveedores", Properties.Resources.cart_fill20, maes_proveed_Click);
            menuStrip1.Items.Add("Tip.Cambio", Properties.Resources.euro20, maes_tipcam_Click);
            //
            menuStrip1.Visible = true;
        }
        private void maes_proveed_Click(object sender, EventArgs e)
        {
            provee ffe1 = new provee();
            ffe1.TopLevel = false;
            ffe1.Parent = this;
            ffe1.Left = pn_pver.Left + pn_pver.Width + 1;
            pn_centro.Controls.Add(ffe1);
            if (this.Width < ffe1.Right + ffe1.Left) this.Width = ffe1.Right + ffe1.Left;
            ffe1.Show();
        }
        private void maes_tipcam_Click(object sender, EventArgs e)
        {
            tipcam ffe1 = new tipcam();
            ffe1.TopLevel = false;
            ffe1.Parent = this;
            ffe1.Left = pn_pver.Left + pn_pver.Width + 1;
            //ffe1.Top = ffe1.Height + pn_centro.Bottom;
            pn_centro.Controls.Add(ffe1);
            if (this.Width < ffe1.Right + ffe1.Left) this.Width = ffe1.Right + ffe1.Left;
            ffe1.Show();
        }
        #endregion

        #region configuración
        private void bt_pcontrol_Click(object sender, EventArgs e)
        {
            pic_icon_menu.Image = Properties.Resources.settings20;
            menuStrip1.Items.Clear();
            menuStrip1.Items.Add("Usuarios", Properties.Resources.happy_20, pcon_users_Click);
            menuStrip1.Items.Add("Permisos", Properties.Resources.permisos20, pcon_permisos_Click);
            menuStrip1.Items.Add("Enlaces", Properties.Resources.link20, pcon_enlaces_Click);
            menuStrip1.Items.Add("Definiciones", Properties.Resources.link20, pcon_definic_Click);
            //
            menuStrip1.Visible = true;
        }
        private void pcon_users_Click(object sender, EventArgs e)
        {
            users enl = new users();
            enl.TopLevel = false;
            enl.Parent = this;
            enl.Left = pn_centro.Width / 3;
            enl.Top = pn_centro.Height / 3;
            pn_centro.Controls.Add(enl);
            enl.Show();
        }
        private void pcon_permisos_Click(object sender, EventArgs e)
        {
            permisos enl = new permisos();
            enl.TopLevel = false;
            enl.Parent = this;
            enl.Left = pn_centro.Width / 3;
            enl.Top = pn_centro.Height / 3;
            pn_centro.Controls.Add(enl);
            enl.Show();
        }
        private void pcon_enlaces_Click(object sender, EventArgs e)
        {
            enlaces enl = new enlaces();
            enl.TopLevel = false;
            enl.Parent = this;
            enl.Left = pn_centro.Width / 3;
            enl.Top = pn_centro.Height / 3;
            pn_centro.Controls.Add(enl);
            enl.Show();
        }
        private void pcon_definic_Click(object sender, EventArgs e)
        {
            defs enl = new defs();
            enl.TopLevel = false;
            enl.Parent = this;
            enl.Left = pn_centro.Width / 3;
            enl.Top = pn_centro.Height / 3;
            pn_centro.Controls.Add(enl);
            enl.Show();
        }
        #endregion

        private void bt_salir_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #endregion

        private void Main_FormClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var aa = MessageBox.Show("Realmente desea salir del sistema?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (aa == DialogResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        #region esta parte es para acortar el acnho del panel de modulos
        private void panelX1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isResizeMode = false;
            }
        }

        private void panelX1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isResizeMode)
            {
                this.Size = new Size(e.X, e.Y);
            }
        }

        private void panelX1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isResizeMode = true;
            }
        }

        private void pn_pver_DoubleClick(object sender, EventArgs e)
        {
            if (pictureBox1.Size.Width == 50)
            {
                pictureBox1.Size = new Size(188, 128);
                bt_finan.Size = new Size(185, 50);
                bt_finan.Text = "Finanzas";
                bt_maestras.Size = new Size(185, 50);
                bt_maestras.Text = "Maestras";
                bt_pcontrol.Size = new Size(185, 50);
                bt_pcontrol.Text = "Configuración";
                bt_salir.Size = new Size(185, 50);
                bt_salir.Text = "Salir";
                bt_serWeb.Size = new Size(185, 50);
                bt_serWeb.Text = "Servicios Web";
                bt_validac.Size = new Size(185, 50);
                bt_validac.Text = "Validación";
                pn_centro.Location = new Point(190, 54);
                pn_menu.Location = new Point(227, 24);
                pn_user.Location = new Point(188, 0);
                pic_icon_menu.Location = new Point(190, 25);
            }
            else
            {
                pictureBox1.Size = new Size(50, 128);
                bt_finan.Size = new Size(50, 50);
                bt_finan.Text = ""; 
                bt_maestras.Size = new Size(50, 50);
                bt_maestras.Text = "";
                bt_pcontrol.Size = new Size(50, 50);
                bt_pcontrol.Text = "";
                bt_salir.Size = new Size(50, 50);
                bt_salir.Text = "";
                bt_serWeb.Size = new Size(50, 50);
                bt_serWeb.Text = "";
                bt_validac.Size = new Size(50, 50);
                bt_validac.Text = "";
                //
                Size aa = pn_centro.Size;
                pn_centro.Location = new Point(55, 54);
                pn_centro.Size = new Size(aa.Width + 135, aa.Height);
                pn_menu.Location = new Point(55, 24);
                pn_user.Location = new Point(50, 0);
                pic_icon_menu.Location = new Point(55, 25);
            }
        }
        #endregion

        private void Main_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.V) bt_validac.PerformClick();
            if (pic_icon_menu.Tag != null && pic_icon_menu.Tag.ToString() != "")
            {
                if (pic_icon_menu.Tag.ToString() == "VALIDACION" &&
                    Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.R) vali_ingresos_Click(null, null);

            }

        }
    }
}
