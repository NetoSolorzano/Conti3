using System;
//using System.Configuration;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class permisos : Form
    {
        static string nomform = "permisos"; // nombre del formulario
        string asd = Conti3.Program.vg_user;   // usuario conectado al sistema
        string colback = "#c0ccea";   // color de fondo del form
        string colpage = "#687086";   // color de los pageframes
        string colgrid = "#a9acb6";   // color fondo sin grillas
        string colstrp = "#687eb6";   // color fuerte
        static string nomtab = "permisos";
        public int totfilgrid, cta;      // variables para impresion
        public string perAg = "";
        public string perMo = "";
        public string perAn = "";
        public string perIm = "";
        libreria lib = new libreria();
        Finan_Egres OFegre = new Finan_Egres();
        // string de conexion
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        DataTable dtg = new DataTable();
        DataTable dtu = new DataTable();

        public permisos()
        {
            InitializeComponent();
        }
        private void permisos_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N) Bt_add.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E) Bt_edit.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P) Bt_print.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.A) Bt_anul.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O) Bt_ver.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S) Bt_close.PerformClick();
        }
        private void permisos_Load(object sender, EventArgs e)
        {
            ToolTip toolTipNombre = new ToolTip();           // Create the ToolTip and associate with the Form container.
            // Set up the delays for the ToolTip.
            toolTipNombre.AutoPopDelay = 5000;
            toolTipNombre.InitialDelay = 1000;
            toolTipNombre.ReshowDelay = 500;
            toolTipNombre.ShowAlways = true;                 // Force the ToolTip text to be displayed whether or not the form is active.
            toolTipNombre.SetToolTip(toolStrip1, nomform);   // Set up the ToolTip text for the object
            init();
            toolboton();
            limpiar(this);
            sololee(this);
            dataload();
            grilla();
            this.KeyPreview = true;
            OFegre.colorea(this, colstrp, colpage, colback);    // pinta el mundo de colores!
            toolStrip1.BackColor = ColorTranslator.FromHtml(colpage);
            // color de boton Bt_graba
            button1.BackColor = ColorTranslator.FromHtml("#667d97");   //  "#656d77",   #f5510f", "#e76433"
            button1.Image = null;
            Bt_add.Enabled = false;
            Bt_anul.Enabled = false;
            comboBox1.SelectedIndex = -1;
            tabControl1.SelectedTab = tabgrilla;
            advancedDataGridView1.Enabled = false;
        }
        private void init()
        {
            Font tiplg = new Font("Arial", 9, FontStyle.Bold);
            advancedDataGridView1.Font = tiplg;
            advancedDataGridView1.DefaultCellStyle.Font = tiplg;
            advancedDataGridView1.RowTemplate.Height = 15;
            advancedDataGridView1.DefaultCellStyle.BackColor = Color.CadetBlue; // Color.CadetBlue;
            jalainfo();
            Bt_add.Image = Properties.Resources.new_tab20; // Image.FromFile(img_btN);
            Bt_edit.Image = Properties.Resources.edit20;
            Bt_anul.Image = Properties.Resources.delete20;
            Bt_close.Image = Properties.Resources.close20;
            Bt_ini.Image = Properties.Resources.arrow_in_left20;
            Bt_sig.Image = Properties.Resources.arrow_right20;
            Bt_ret.Image = Properties.Resources.arrow_left20;
            Bt_fin.Image = Properties.Resources.arrow_in_right20;
        }
        private void grilla()                   // arma la grilla
        {
            // id,rutaf,formulario,btn1,btn2,btn3,btn4,btn5,btn6,btn7,btn8,coment,usuario
            Font tiplg = new Font("Arial",7, FontStyle.Bold);
            advancedDataGridView1.Font = tiplg;
            advancedDataGridView1.DefaultCellStyle.Font = tiplg;
            advancedDataGridView1.RowTemplate.Height = 15;
            advancedDataGridView1.DefaultCellStyle.BackColor = Color.CadetBlue; // Color.MediumAquamarine;
            advancedDataGridView1.DataSource = dtg;
            // id 
            advancedDataGridView1.Columns[0].Visible = false;
            // ruta menu
            advancedDataGridView1.Columns[1].Visible = true;            // columna visible o no
            advancedDataGridView1.Columns[1].HeaderText = "RUTA DE MENU";    // titulo de la columna
            advancedDataGridView1.Columns[1].Width = 100;                // ancho
            advancedDataGridView1.Columns[1].ReadOnly = true;           // lectura o no
            advancedDataGridView1.Columns[1].Tag = "validaNO";
            advancedDataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // formulario
            advancedDataGridView1.Columns[2].Visible = false;
            // btn1
            advancedDataGridView1.Columns[3].Visible = true;       
            advancedDataGridView1.Columns[3].HeaderText = "NUEVO";
            advancedDataGridView1.Columns[3].Width = 50;
            advancedDataGridView1.Columns[3].ReadOnly = false;          // las celdas de esta columna pueden cambiarse
            advancedDataGridView1.Columns[3].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn2
            advancedDataGridView1.Columns[4].Visible = true;
            advancedDataGridView1.Columns[4].HeaderText = "EDITA";
            advancedDataGridView1.Columns[4].Width = 50;
            advancedDataGridView1.Columns[4].ReadOnly = false;          // las celdas de esta columna pueden cambiarse
            advancedDataGridView1.Columns[4].Tag = "validaSI";          // las celdas de esta columna se validan
            advancedDataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn3
            advancedDataGridView1.Columns[5].Visible = true;       
            advancedDataGridView1.Columns[5].HeaderText = "ANULA";
            advancedDataGridView1.Columns[5].Width = 50;
            advancedDataGridView1.Columns[5].ReadOnly = false;
            advancedDataGridView1.Columns[5].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn4
            advancedDataGridView1.Columns[6].Visible = true;       
            advancedDataGridView1.Columns[6].HeaderText = "VER";
            advancedDataGridView1.Columns[6].Width = 50;
            advancedDataGridView1.Columns[6].ReadOnly = false;
            advancedDataGridView1.Columns[6].Tag = "validaSI";          // las celdas de esta columna SI se validan
            advancedDataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn5
            advancedDataGridView1.Columns[7].Visible = true;
            advancedDataGridView1.Columns[7].HeaderText = "IMPRIME";
            advancedDataGridView1.Columns[7].Width = 50;
            advancedDataGridView1.Columns[7].ReadOnly = false;
            advancedDataGridView1.Columns[7].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn6
            advancedDataGridView1.Columns[8].Visible = true;
            advancedDataGridView1.Columns[8].HeaderText = "SALIR";
            advancedDataGridView1.Columns[8].Width = 50;
            advancedDataGridView1.Columns[8].ReadOnly = true;
            advancedDataGridView1.Columns[8].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn7
            advancedDataGridView1.Columns[9].Visible = true;
            advancedDataGridView1.Columns[9].HeaderText = "PREVIEW";
            advancedDataGridView1.Columns[9].Width = 50;
            advancedDataGridView1.Columns[9].ReadOnly = false;
            advancedDataGridView1.Columns[9].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // btn8
            advancedDataGridView1.Columns[10].Visible = true;
            advancedDataGridView1.Columns[10].HeaderText = "EXPORTA";
            advancedDataGridView1.Columns[10].Width = 50;
            advancedDataGridView1.Columns[10].ReadOnly = false;
            advancedDataGridView1.Columns[10].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[10].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // coment
            advancedDataGridView1.Columns[11].Visible = true;
            advancedDataGridView1.Columns[11].HeaderText = "COMENTARIOS";
            advancedDataGridView1.Columns[11].Width = 150;
            advancedDataGridView1.Columns[11].ReadOnly = true;
            advancedDataGridView1.Columns[11].Tag = "validaNO";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //usuario
            advancedDataGridView1.Columns[12].Visible = true;
            advancedDataGridView1.Columns[12].HeaderText = "USUARIO";
            advancedDataGridView1.Columns[12].Width = 70;
            advancedDataGridView1.Columns[12].ReadOnly = true;
            advancedDataGridView1.Columns[12].Tag = "validaSI";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[12].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }
        private void jalainfo()                 // obtiene datos de imagenes
        {
            try
            {
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                string consulta = "select campo,param,valor from enlaces where formulario=@nofo";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@nofo", "main");   // nomform
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                for (int t = 0; t < dt.Rows.Count; t++)
                {
                    DataRow row = dt.Rows[t];
                    if (row["campo"].ToString() == "imagenes")
                    {
                        // antes había acá las rutas de los iconos
                    }
                }
                da.Dispose();
                dt.Dispose();
                conn.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error de conexión");
                Application.Exit();
                return;
            }
        }
        public void jalaoc(string campo)        // jala datos de usuarios por id o nom_user
        {
            if (campo == "tx_idr" && tx_idr.Text != "")
            {
                //SELECT id, rutaf, formulario, btn1, btn2, btn3, btn4, btn5, btn6, coment, usuario FROM coop2018.permisos; textBox1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[1].Value.ToString();  // formulario
                textBox1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[1].Value.ToString();  // ruta
                textBox2.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[12].Value.ToString();  // usuario
                textBox3.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[11].Value.ToString();  // comentario
                chk_nuevo.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[3].Value.ToString() == "S") ? true : false;
                chk_edita.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[4].Value.ToString() == "S") ? true : false;
                chk_anula.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[5].Value.ToString() == "S") ? true : false;
                chk_visua.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[6].Value.ToString() == "S") ? true : false;
                chk_impri.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[7].Value.ToString() == "S") ? true : false;
                chk_salir.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[8].Value.ToString() == "S") ? true : false;
                chk_prev.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[9].Value.ToString() == "S") ? true : false;
                chk_expor.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[10].Value.ToString() == "S") ? true : false;
                comboBox1.SelectedValue = textBox2.Text;
            }
        }
        public void dataload()                  // jala datos para los combos y la grilla
        {
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State != ConnectionState.Open)
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
                return;
            }
            tabControl1.SelectedTab = tabreg;
            // datos de los permisos
            string datgri = "SELECT id,rutaf,formulario,btn1,btn2,btn3,btn4,btn5,btn6,btn7,btn8,coment,usuario FROM permisos order by usuario,rutaf";
            MySqlCommand cdg = new MySqlCommand(datgri, conn);
            MySqlDataAdapter dag = new MySqlDataAdapter(cdg);
            dtg.Clear();
            dag.Fill(dtg);
            dag.Dispose();
            //  datos para el combobox de usaurios
            this.comboBox1.Items.Clear();
            string datuse = "select nom_user,nom_user from usuarios where bloqueado=@bloq";
            MySqlCommand cdu = new MySqlCommand(datuse, conn);
            cdu.Parameters.AddWithValue("@bloq", 0);
            MySqlDataAdapter dacu = new MySqlDataAdapter(cdu);
            dtu.Clear();
            dacu.Fill(dtu);
            comboBox1.DataSource = dtu;
            comboBox1.DisplayMember = "nom_user";
            comboBox1.ValueMember = "nom_user";
            //
            dacu.Dispose();
            conn.Close();
        }
        string[] equivinter(string titulo)        // equivalencia entre titulo de columna y tabla 
        {
            string[] retorna = new string[2];
            switch (titulo)
            {
                case "NIVEL":
                    retorna[0] = "desc_niv";
                    retorna[1] = "codigo";
                    break;
                case "???":
                    retorna[0] = "";
                    retorna[1] = "";
                    break;
                case "????":
                    retorna[0] = "";
                    retorna[1] = "";
                    break;
                case "LOCAL":
                    retorna[0] = "desc_alm";
                    retorna[1] = "idcodice";
                    break;
                case "TIENDA":
                    retorna[0] = "desc_ven";
                    retorna[1] = "idcodice";
                    break;
                case "SEDE":
                    retorna[0] = "desc_loc";
                    retorna[1] = "idcodice";
                    break;
                case "RUC":
                    retorna[0] = "desc_raz";
                    retorna[1] = "idcodice";
                    break;
            }
            return retorna;
        }

        #region limpiadores_modos
        public void sololee(Form lfrm)
        {
            foreach (Control oControls in lfrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is ComboBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is RadioButton)
                {
                    oControls.Enabled = false;
                }
                if (oControls is DateTimePicker)
                {
                    oControls.Enabled = false;
                }
                if (oControls is MaskedTextBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is GroupBox)
                {
                    oControls.Enabled = false;
                }
            }
        }
        public void escribe(Form efrm)
        {
            foreach (Control oControls in efrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is ComboBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is RadioButton)
                {
                    oControls.Enabled = true;
                }
                if (oControls is DateTimePicker)
                {
                    oControls.Enabled = true;
                }
                if (oControls is MaskedTextBox)
                {
                    oControls.Enabled = true;
                }
            }
        }
        public static void limpiar(Form ofrm)
        {
            foreach (Control oControls in ofrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
            }
        }
        public void limpiapag(TabPage pag)
        {
            tabControl1.SelectedTab = pag;
            foreach (Control oControls in pag.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
                if(oControls is CheckBox)
                {
                    chk_anula.Checked = false;
                    chk_edita.Checked = false;
                    chk_expor.Checked = false;
                    chk_impri.Checked = false;
                    chk_nuevo.Checked = false;
                    chk_prev.Checked = false;
                    chk_salir.Checked = false;
                    chk_visua.Checked = false;
                }
                if(oControls is ComboBox)
                {
                    comboBox1.SelectedIndex = -1;
                }
            }
        }
        public void limpia_chk()    
        {
            checkBox1.Checked = false;
            chk_anula.Checked = false;
            chk_edita.Checked = false;
            chk_nuevo.Checked = false;
            chk_visua.Checked = false;
            chk_impri.Checked = false;
            chk_salir.Checked = false;
            chk_prev.Checked = false;
            chk_expor.Checked = false;
        }
        public void limpia_otros()
        {
            //this.checkBox1.Checked = false;
        }
        public void limpia_combos()
        {
            //this.comboBox1.SelectedIndex = -1;
        }
        #endregion limpiadores_modos;

        #region boton_form GRABA EDITA ANULA
        private void button1_Click(object sender, EventArgs e)
        {
            // validamos que los campos no esten vacíos
            if (textBox1.Text == "")
            {
                MessageBox.Show("Ingrese la ruta del menu", " Error! ");
                textBox1.Focus();
                return;
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("Ingrese el usuario", " Error! ");
                textBox2.Focus();
                return;
            }
            // grabamos, actualizamos, etc
            string modo = Tx_modo.Text;
            string iserror = "no";
            string asd = Conti3.Program.vg_user;
            string verapp = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
            if (modo == "NUEVO")
            {
                // no existe nuevo
                // SELECT id,rutaf,formulario,btn1,btn2,btn3,btn4,btn5,btn6,coment,usuario FROM coop2018.permisos;
            }
            if (modo == "EDITAR")
            {
                string consulta = "update permisos set " +
                        "rutaf=@rut,usuario=@usu,btn1=@bt1,btn2=@bt2,btn3=@bt3,btn4=@bt4,btn5=@bt5,btn6=@bt6,btn7=@bt7,btn8=@bt8," +
                        "coment=@com,verApp=@vap,userm=@asd,fechm=now(),diriplan4=@ipl,diripwan4=@ipw,netbname=@nbna " +
                        "where id=@idc";
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    MySqlCommand mycom = new MySqlCommand(consulta, conn);
                    mycom.Parameters.AddWithValue("@idc", tx_idr.Text);
                    mycom.Parameters.AddWithValue("@rut", textBox1.Text);
                    mycom.Parameters.AddWithValue("@usu", textBox2.Text);
                    mycom.Parameters.AddWithValue("@bt1", (chk_nuevo.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt2", (chk_edita.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt3", (chk_anula.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt4", (chk_visua.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt5", (chk_impri.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt6", (chk_salir.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt7", (chk_prev.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@bt8", (chk_expor.Checked == true) ? "S" : "N");
                    mycom.Parameters.AddWithValue("@com", textBox3.Text);
                    mycom.Parameters.AddWithValue("@vap", verapp);
                    mycom.Parameters.AddWithValue("@asd", asd);
                    mycom.Parameters.AddWithValue("@ipl", lib.iplan());
                    mycom.Parameters.AddWithValue("@ipw", Conti3.Program.vg_ipwan);
                    mycom.Parameters.AddWithValue("@nbna", Environment.MachineName);
                    try
                    {
                        mycom.ExecuteNonQuery();
                        // actualizamos el datatable
                        for (int i = 0; i < dtg.Rows.Count; i++)
                        {
                            DataRow row = dtg.Rows[i];
                            if (row[0].ToString() == tx_idr.Text)
                            {
                                // id,rutaf,formulario,btn1,btn2,btn3,btn4,btn5,btn6,btn7,btn8,coment,usuario
                                dtg.Rows[i][1] = textBox1.Text.Trim();
                                dtg.Rows[i][3] = (chk_nuevo.Checked == true) ? "S" : "N";
                                dtg.Rows[i][4] = (chk_edita.Checked == true) ? "S" : "N";
                                dtg.Rows[i][5] = (chk_anula.Checked == true) ? "S" : "N";
                                dtg.Rows[i][6] = (chk_visua.Checked == true) ? "S" : "N";
                                dtg.Rows[i][7] = (chk_impri.Checked == true) ? "S" : "N";
                                dtg.Rows[i][8] = (chk_salir.Checked == true) ? "S" : "N";
                                dtg.Rows[i][9] = (chk_prev.Checked == true) ? "S" : "N";
                                dtg.Rows[i][10] = (chk_expor.Checked == true) ? "S" : "N";
                                dtg.Rows[i][11] = textBox3.Text.Trim();
                            }
                        }
                        // actualizamos la tabla seguimiento de usuarios
                        string resulta = lib.ult_mov(nomform, nomtab, asd);
                        if (resulta != "OK")
                        {
                            MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Exit();
                            return;
                        }
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error de Editar definición",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        iserror = "si";
                    }
                    conn.Close();
                    //permisos();
                }
                else
                {
                    MessageBox.Show("No se estableció conexión con el servidor", "Atención - no se puede continuar");
                    Application.Exit();
                    return;
                }
            }
            if (modo == "ANULAR")       // opción para borrar
            { 
                // no se anulan, solo se habilitan o deshabilitan
            }
            if (iserror == "no")
            {
                // debe limpiar los campos y actualizar la grilla
                limpiar(this);
                limpiapag(tabreg);
                limpia_otros();
                this.textBox1.Focus();
                //dataload();
            }
        }
        #endregion boton_form;

        #region leaves
        private void tx_idr_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "NUEVO" && tx_idr.Text != "")
            {
                //string aca = tx_idr.Text;
                //limpia_chk();
                //limpia_combos();
                //limpiar(this);
                //tx_idr.Text = aca;
                jalaoc("tx_idr");               // jalamos los datos del registro
            }
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {
            /*  validamos segun el modo
            */
        }
        #endregion leaves;

        #region botones_de_comando_y_permisos  
        public void toolboton()
        {
            DataTable mdtb = new DataTable();
            const string consbot = "select * from permisos where formulario=@nomform and usuario=@use";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    MySqlCommand consulb = new MySqlCommand(consbot, conn);
                    consulb.Parameters.AddWithValue("@nomform", nomform);
                    consulb.Parameters.AddWithValue("@use", asd);
                    MySqlDataAdapter mab = new MySqlDataAdapter(consulb);
                    mab.Fill(mdtb);
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, " Error ");
                    return;
                }
                finally { conn.Close(); }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
                return;
            }
            if (mdtb.Rows.Count > 0)
            {
                DataRow row = mdtb.Rows[0];
                if (Convert.ToString(row["btn1"]) == "S")               // nuevo
                {
                    this.Bt_add.Visible = true;
                }
                else { this.Bt_add.Visible = false; }
                if (Convert.ToString(row["btn2"]) == "S")               // editar
                {
                    this.Bt_edit.Visible = true;
                }
                else { this.Bt_edit.Visible = false; }
                if (Convert.ToString(row["btn3"]) == "S")               // anular
                {
                    this.Bt_anul.Visible = true;
                }
                else { this.Bt_anul.Visible = false; }
                if (Convert.ToString(row["btn4"]) == "S")               // visualizar
                {
                    
                }
                //else { this.Bt_anul.Visible = false; }
                if (Convert.ToString(row["btn5"]) == "S")               // imprimir
                {
                    
                }
                //else { //this.bt_print.Visible = false; }
                if (Convert.ToString(row["btn6"]) == "S")               // salir del form
                {
                    this.Bt_close.Visible = true;
                }
                else { this.Bt_close.Visible = false; }
                if (Convert.ToString(row["btn7"]) == "S")               // vista preliminar
                {
                    //this.bt_view.Visible = true;
                }
                //else { //this.bt_view.Visible = false; }
                if (Convert.ToString(row["btn8"]) == "S")               // exportar xls
                {
                    //this.bt_view.Visible = true;
                }
                //else { //this.bt_view.Visible = false; }
            }
            else
            {
                Bt_add.Visible = false;
                Bt_edit.Visible = false;
                Bt_anul.Visible = false;
                Bt_close.Visible = true;
            }
        }
        #region botones
        private void Bt_add_Click(object sender, EventArgs e)
        {
            /*  advancedDataGridView1.Enabled = true;
            tabControl1.SelectedTab = tabreg;
            escribe(this);
            Tx_modo.Text = "NUEVO";
            button1.Image = Conti3.Properties.Resources.save_negro40;
            textBox1.Focus();
            limpiar(this);
            limpia_otros();
            limpia_combos();    */
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            advancedDataGridView1.Enabled = true;
            //string codu = "";
            string idr = "";
            if (advancedDataGridView1.CurrentRow.Index > -1)
            {
                //codu = advancedDataGridView1.CurrentRow.Cells[1].Value.ToString();
                idr = advancedDataGridView1.CurrentRow.Cells[0].Value.ToString();
                tx_rind.Text = advancedDataGridView1.CurrentRow.Index.ToString();
            }
            tabControl1.SelectedTab = tabgrilla;
            escribe(this);
            Tx_modo.Text = "EDITAR";
            button1.Image = Conti3.Properties.Resources.save_negro40;
            limpiar(this);
            limpia_otros();
            limpia_combos();
            jalaoc("tx_idr");
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            sololee(this);
            this.Tx_modo.Text = "IMPRIMIR";
            this.button1.Image = Conti3.Properties.Resources.printer40;
            this.textBox1.Focus();
        }
        private void Bt_anul_Click(object sender, EventArgs e)          // pone todos los permisos en N
        {
            /*
            advancedDataGridView1.Enabled = true;
            string codu = "";
            string idr = "";
            if (advancedDataGridView1.CurrentRow.Index > -1)
            {
                codu = advancedDataGridView1.CurrentRow.Cells[1].Value.ToString();
                idr = advancedDataGridView1.CurrentRow.Cells[0].Value.ToString();
                tx_rind.Text = advancedDataGridView1.CurrentRow.Index.ToString();
            }
            tabControl1.SelectedTab = tabreg;
            escribe(this);
            Tx_modo.Text = "ANULAR";
            button1.Image = Conti3.Properties.Resources.delete40;
            limpiar(this);
            limpia_otros();
            limpia_combos();
            jalaoc("tx_idr");
            */
        }
        private void Bt_first_Click(object sender, EventArgs e)
        {
            limpiar(this);
            limpia_chk();
            limpia_combos();
            //--
            tx_idr.Text = lib.gofirts(nomtab);
            tx_idr_Leave(null, null);
        }
        private void Bt_back_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpia_combos();
            limpiar(this);
            //--
            tx_idr.Text = lib.goback(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_next_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpia_combos();
            limpiar(this);
            //--
            tx_idr.Text = lib.gonext(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_last_Click(object sender, EventArgs e)
        {
            limpiar(this);
            limpia_chk();
            limpia_combos();
            //--
            tx_idr.Text = lib.golast(nomtab);
            tx_idr_Leave(null, null);
        }
        #endregion botones;
        // permisos para habilitar los botones de comando
        #endregion botones_de_comando  ;

        #region comboboxes
        // selected index del combobox de usuarios
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > -1)
            {
                DataRow row = ((DataTable)comboBox1.DataSource).Rows[comboBox1.SelectedIndex];
                textBox2.Text = (string)row["nom_user"];
                //int Id = (int)row["idcodice"];
            }
            else
            {
                textBox2.Text = "";
            }
        }
        #endregion comboboxes

        #region advancedatagridview
        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)                  // filtro de las columnas
        {
            dtg.DefaultView.RowFilter = advancedDataGridView1.FilterString;
        }
        private void advancedDataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)            // almacena valor previo al ingresar a la celda
        {
            advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim();
        }
        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex == 1)
            {
                //string codu = "";
                string idr = "";
                idr = advancedDataGridView1.CurrentRow.Cells[0].Value.ToString();
                tx_rind.Text = advancedDataGridView1.CurrentRow.Index.ToString();
                tabControl1.SelectedTab = tabreg;
                limpiar(this);
                limpia_otros();
                limpia_combos();
                tx_idr.Text = idr;
                jalaoc("tx_idr");
            }
        }
        private void button2_Click(object sender, EventArgs e)  // carga a grilla permisos del usuario
        {
            // nos conectamos a la bd
            // seleccionamos los permisos del usuario
            // pasamos los datos al adgv

        }
        private void advancedDataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e) // valida cambios en valor de la celda
        {
            if (e.RowIndex > -1 && e.ColumnIndex > 0 
                && advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() != e.FormattedValue.ToString())
            {
                string campo = advancedDataGridView1.Columns[e.ColumnIndex].Name.ToString();
                string[] noeta = equivinter(advancedDataGridView1.Columns[e.ColumnIndex].HeaderText.ToString());    // retorna la tabla segun el titulo de la columna

                var aaa = MessageBox.Show("Confirma que desea cambiar el valor?",
                    "Columna: " + advancedDataGridView1.Columns[e.ColumnIndex].HeaderText.ToString(),
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    if(advancedDataGridView1.Columns[e.ColumnIndex].Tag.ToString() == "validaSI")   // la columna se valida?
                    {
                        // valida si el dato ingresado es valido en la columna
                        if(e.ColumnIndex == 12)                         // valida usuarios
                        {
                            if (lib.validac("usuarios", "nom_user", e.FormattedValue.ToString()) == true)
                            {
                                // llama a libreria con los datos para el update - tabla,id,campo,nuevo valor
                                lib.actuac(nomtab, campo, e.FormattedValue.ToString(), advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                            }
                            else
                            {
                                MessageBox.Show("El valor no es válido para la columna", "Atención - Corrija");
                                e.Cancel = true;
                            }
                        }
                        if(e.ColumnIndex > 2 && e.ColumnIndex < 11)  // valida S ó N
                        {
                            if(e.FormattedValue.ToString().Trim() != "S" && e.FormattedValue.ToString().Trim() != "N")
                            {
                                MessageBox.Show("El valor no es válido para la columna", "Atención - Corrija");
                                e.Cancel = true;
                            }
                            else
                            {
                                // llama a libreria con los datos para el update - tabla,id,campo,nuevo valor
                                lib.actuac(nomtab, campo, e.FormattedValue.ToString(), advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                            }
                        }
                    }
                    else
                    {
                        // llama a libreria con los datos para el update - tabla,id,campo,nuevo valor
                        lib.actuac(nomtab, campo, e.FormattedValue.ToString(), advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                    }
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }
        #endregion
    }
}
