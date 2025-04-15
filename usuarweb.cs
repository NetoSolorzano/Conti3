using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class usuarweb : Form
    {
        static string nomform = "usuarweb"; // nombre del formulario
        string asd = Conti3.Program.vg_user;   // usuario conectado al sistema
        string verapp = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
        static string nomtab = "webuser";
        public int totfilgrid, cta;      // variables para impresion
        public string perAg = "";
        public string perMo = "";
        public string perAn = "";
        public string perIm = "";
        libreria lib = new libreria();
        Finan_Egres OFegre = new Finan_Egres();
        ccolores OColores = new ccolores();
        cajDestino Ocajd = new cajDestino();
        // string de conexion
        //static string serv = ConfigurationManager.AppSettings["serv"].ToString();
        static string port = ConfigurationManager.AppSettings["port"].ToString();
        //static string usua = ConfigurationManager.AppSettings["user"].ToString();
        //static string cont = ConfigurationManager.AppSettings["pass"].ToString();
        static string data = ConfigurationManager.AppSettings["data"].ToString();
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + data + ";";
        DataTable dtg = new DataTable();
        //DataTable dt_ctaPer = new DataTable();
        List<string> lista_ = new List<string>();                                   // cuentas personales 

        public usuarweb()
        {
            InitializeComponent();
            //dt_ctaPer.Columns.Add("idcodice");
            //dt_ctaPer.Columns.Add("descrizione");
        }
        private void usuarweb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N) Bt_add.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E) Bt_edit.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P) Bt_print.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.A) Bt_anul.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O) Bt_ver.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S) Bt_close.PerformClick();
        }
        private void usuarweb_Load(object sender, EventArgs e)
        {
            /*
            ToolTip toolTipNombre = new ToolTip();           // Create the ToolTip and associate with the Form container.
            // Set up the delays for the ToolTip.
            toolTipNombre.AutoPopDelay = 5000;
            toolTipNombre.InitialDelay = 1000;
            toolTipNombre.ReshowDelay = 500;
            toolTipNombre.ShowAlways = true;                 // Force the ToolTip text to be displayed whether or not the form is active.
            toolTipNombre.SetToolTip(toolStrip1, nomform);   // Set up the ToolTip text for the object
            */
            init();
            limpiar(this);
            sololee(this);
            this.KeyPreview = true;
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                toolboton(conn);
                dataload(conn);
                OFegre.jalacolores(conn, OColores, nomform);
            }
            OFegre.colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);
            //toolStrip1.BackColor = ColorTranslator.FromHtml(OColores.Fondo_pageFrame);   // colpage
            tabgrilla.BackColor = ColorTranslator.FromHtml(OColores.Fondo_pageFrame);
            tabuser.BackColor = ColorTranslator.FromHtml(OColores.Fondo_pageFrame);
            button1.BackColor = ColorTranslator.FromHtml(OColores.Fondo_boton_graba);   // "#667d97"
            button1.Image = null;
            tabControl1.SelectedTab = tabgrilla;
            advancedDataGridView1.Enabled = false;
            grilla();
        }
        private void init()
        {
            tx_codigo.CharacterCasing = CharacterCasing.Upper;
            tx_codigo.MaxLength = 10;
            tx_pass.MaxLength = 10;
            tx_usweb.CharacterCasing = CharacterCasing.Upper;
            tx_usweb.MaxLength = 50;
            Tx_ctaDes.CharacterCasing = CharacterCasing.Upper;
            jalainfo();
            Bt_add.Image = Properties.Resources.new_tab20;
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
            Font tiplg = new Font("Arial",7, FontStyle.Bold);
            advancedDataGridView1.Font = tiplg;
            advancedDataGridView1.DefaultCellStyle.Font = tiplg;
            advancedDataGridView1.RowTemplate.Height = 15;
            advancedDataGridView1.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(OColores.Grilla_fila_normal);    // Color.CadetBlue
            advancedDataGridView1.DataSource = dtg;
            // a.code,a.names,a.passw,a.central,a.count,b.Descrizione,a.blocked
            if (dtg.Rows.Count > 1)
            {
                for (int i = 0; i < dtg.Columns.Count; i++)
                {
                    {
                        advancedDataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        _ = decimal.TryParse(advancedDataGridView1.Rows[0].Cells[i].Value.ToString(), out decimal vd);
                        if (vd != 0) advancedDataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                }
                int b = 0;
                for (int i = 0; i < dtg.Columns.Count; i++)
                {
                    int a = advancedDataGridView1.Columns[i].Width;
                    b += a;
                    advancedDataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    advancedDataGridView1.Columns[i].Width = a;
                }
            }
            // Código del usuario
            advancedDataGridView1.Columns["code"].Visible = true;
            advancedDataGridView1.Columns["code"].HeaderText = "CODIGO";
            advancedDataGridView1.Columns["code"].ReadOnly = true;
            // nombre del usuario
            advancedDataGridView1.Columns["names"].Visible = true;            // columna visible o no
            advancedDataGridView1.Columns["names"].HeaderText = "USUARIO";    // titulo de la columna
            advancedDataGridView1.Columns["names"].ReadOnly = true;           // lectura o no
            // passw
            advancedDataGridView1.Columns["passw"].Visible = false;
            // 
            advancedDataGridView1.Columns["central"].Visible = false;       
            // cuenta personal
            advancedDataGridView1.Columns["count"].Visible = true;
            advancedDataGridView1.Columns["count"].HeaderText = "CUENTA";
            advancedDataGridView1.Columns["count"].ReadOnly = true;
            // nombre de la cuenta
            advancedDataGridView1.Columns["Descrizione"].Visible = true;
            advancedDataGridView1.Columns["Descrizione"].HeaderText = "NOMBRE DE CUENTA";
            advancedDataGridView1.Columns["Descrizione"].ReadOnly = true;
            // bloqueado
            advancedDataGridView1.Columns["blocked"].Visible = true;       
            advancedDataGridView1.Columns["blocked"].HeaderText = "BLOQ";
            advancedDataGridView1.Columns["blocked"].ReadOnly = true;
            //
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
                    // si es necesario ponemos acá
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
            if(tx_rind.Text.Trim() != "")
            {
                // a.code,a.names,a.passw,a.central,a.count,b.Descrizione,a.blocked
                tx_codigo.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells["code"].Value.ToString().Trim();
                tx_pass.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells["passw"].Value.ToString();
                tx_usweb.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells["names"].Value.ToString().Trim();
                tx_dat_cuenta.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells["count"].Value.ToString().Trim();
                if (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells["blocked"].Value.ToString() == "1") checkBox1.Checked = true;
                else checkBox1.Checked = false;
                //                 
                string[] vuelto = OFegre.ValiCtaCon(advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells["Descrizione"].Value.ToString().Trim(), "PER", "algo");
                if (vuelto.Length > 0 && vuelto[0] != "")   // 
                {
                    Ocajd.codigo = vuelto[0];
                    Ocajd.nombre = vuelto[1];
                    Ocajd.largo = vuelto[2];
                    Tx_ctaDes.Text = Ocajd.largo;
                }
                else
                {
                    Tx_ctaDes.Clear();
                    tx_dat_cuenta.Text = "";
                    MessageBox.Show("No existe el nombre de la cuenta");
                }
            }
        }
        public void dataload(MySqlConnection conn)                  // jala datos para los combos y la grilla
        {
            //MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            //conn.Open();
            if (conn.State != ConnectionState.Open)
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
                return;
            }
            tabControl1.SelectedTab = tabuser;
            // datos de combo cuenta personal
            lista_.Clear();
            lista_ = new List<string>();
            DataRow[] _pers = Program.dt_definic.Select("idtabella='CON' and numero='1'", "descrizione ASC");
            foreach (DataRow row in _pers)
            {
                lista_.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_ctaDes.Values = lista_.ToArray();
            // datos de usuarios web
            string datgri = "SELECT a.code,a.names,a.passw,a.central,a.count,b.Descrizione,a.blocked " +
                "from webuser a LEFT JOIN desc_con b ON b.IDCodice=a.count";
            MySqlCommand cdg = new MySqlCommand(datgri, conn);
            MySqlDataAdapter dag = new MySqlDataAdapter(cdg);
            dtg.Clear();
            dag.Fill(dtg);
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
                case "TIPO":
                    retorna[0] = "desc_tpu";
                    retorna[1] = "idcodice";
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
        private void pintaFilaAnul()
        {
            for (int i = 0; i < advancedDataGridView1.Rows.Count - 1; i++)
            {
                if (advancedDataGridView1.Rows[i].Cells[6].Value.ToString() == "1")
                {
                    advancedDataGridView1.Rows[i].DefaultCellStyle.BackColor = ColorTranslator.FromHtml(OColores.Grilla_fila_anulada);
                }
                else 
                { 
                    advancedDataGridView1.Rows[i].DefaultCellStyle.BackColor = ColorTranslator.FromHtml(OColores.Grilla_fila_normal);
                    advancedDataGridView1.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
                }
            }
        }
        private void xxx()
        {
            if (Tx_ctaDes.Text.Trim() != "" && Tx_ctaDes.Text.Length > 3)  // 
            {
                string[] vuelto = OFegre.ValiCtaCon(Tx_ctaDes.Text, "PER", "algo");
                if (vuelto.Length > 0 && vuelto[0] != "")
                {
                    Ocajd.codigo = vuelto[0];
                    Ocajd.nombre = vuelto[1];
                    Ocajd.largo = vuelto[2];
                    //eti_nomCaja.Text = Ocajd.nombre;
                    tx_dat_cuenta.Text = Ocajd.codigo;
                }
                else
                {
                    Tx_ctaDes.Clear();
                    tx_dat_cuenta.Text = "";
                    Ocajd.codigo = "";
                    Ocajd.nombre = "";
                    Ocajd.largo = "";
                    MessageBox.Show("No existe el nombre de la cuenta");
                }
            }
            else
            {
                Ocajd.codigo = "";
                Ocajd.nombre = "";
                Ocajd.largo = "";
                tx_dat_cuenta.Text = "";
            }

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
                if (oControls is CheckBox)
                {
                    oControls.Enabled = false;
                }
            }
        }
        public void sololeePag(TabPage pag)
        {
            foreach (Control oControls in pag.Controls)
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
                if (oControls is CheckBox)
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
                if (oControls is CheckBox)
                {
                    oControls.Enabled = true;
                }
            }
        }
        public void escribePag(TabPage pag)
        {
            foreach (Control oControls in pag.Controls)
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
                if (oControls is GroupBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is CheckBox)
                {
                    oControls.Enabled = true;
                }
            }
        }
        private void limpiar(Form ofrm)   // public static void 
        {
            foreach (Control oControls in ofrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
            }
            Ocajd.codigo = "";
            Ocajd.nombre = "";
            Ocajd.largo = "";
        }
        public void limpiapag(TabPage pag)
        {
            foreach (Control oControls in pag.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
            }
        }
        public void limpia_chk()    
        {
            checkBox1.Checked = false;
        }
        public void limpia_otros(TabPage pag)
        {
            tabControl1.SelectedTab = pag;
            this.checkBox1.Checked = false;
        }
        public void limpia_combos(TabPage pag)
        {
            tabControl1.SelectedTab = pag;
            //cmb_destin.SelectedIndex = -1;
            tx_dat_cuenta.Text = "";
        }
        #endregion limpiadores_modos;

        #region boton_form GRABA EDITA ANULA
        private void button1_Click(object sender, EventArgs e)
        {
            // validamos que los campos no esten vacíos
            if (tx_codigo.Text.Trim() == "")
            {
                MessageBox.Show("El código no puede estar vacío", " Error! ");
                tx_codigo.Focus();
                return;
            }
            if (tx_pass.Text.Trim() == "")
            {
                MessageBox.Show("La contraseña no puede estar vacía", " Error! ");
                tx_pass.Focus();
                return;
            }
            if (tx_usweb.Text.Trim() == "")
            {
                MessageBox.Show("El nombre del usuario no puede estar vacío", " Error! ");
                tx_usweb.Focus();
                return;
            }
            if (tx_dat_cuenta.Text == "")
            {
                MessageBox.Show("Seleccione la cuenta de destino", " Error! ");
                Tx_ctaDes.Focus();
                return;
            }
            if (Tx_ctaDes.Text.Trim() == "")
            {
                MessageBox.Show("Escriba la cuenta de destino", " Error! ");
                Tx_ctaDes.Focus();
                return;
            }
            // grabamos, actualizamos, etc
            string modo = this.Tx_modo.Text;
            string iserror = "no";
            if (modo == "NUEVO")
            {
                var mes = MessageBox.Show("Realmente desea AGREGAR el usuario web?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (mes == DialogResult.Yes)
                {
                    string consulta = "insert into webuser (code,names,passw,central,count,blocked,verApp,userc,fechc,diriplan4,diripwan4,netbname) " +
                        "values (@p5txt1,@p5txt2,@p5txt3,@p5com2,@p5com1,@p5chk1,@verap,@use,now(),@dirl,@dirw,@netb)";
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        MySqlCommand mycomand = new MySqlCommand(consulta, conn);
                        mycomand.Parameters.AddWithValue("@p5txt1", tx_codigo.Text);
                        mycomand.Parameters.AddWithValue("@p5txt2", tx_usweb.Text);
                        mycomand.Parameters.AddWithValue("@p5txt3", tx_pass.Text);
                        mycomand.Parameters.AddWithValue("@p5com2", "LIM");
                        mycomand.Parameters.AddWithValue("@p5com1", tx_dat_cuenta.Text);
                        mycomand.Parameters.AddWithValue("@p5chk1", (checkBox1.Checked == true) ? 1 : 0);
                        mycomand.Parameters.AddWithValue("@verap", verapp);
                        mycomand.Parameters.AddWithValue("@use", asd);
                        mycomand.Parameters.AddWithValue("@dirl", lib.iplan());
                        mycomand.Parameters.AddWithValue("@dirw", Conti3.Program.vg_ipwan);
                        mycomand.Parameters.AddWithValue("@netb", Environment.MachineName);
                        try
                        {
                            mycomand.ExecuteNonQuery();
                            mycomand = new MySqlCommand("select last_insert_id()", conn);
                            MySqlDataReader dr = mycomand.ExecuteReader();
                            string idtu = "";
                            if (dr.Read()) idtu = dr.GetInt32(0).ToString();
                            dr.Close();
                            string resulta = lib.ult_mov(nomform, nomtab, asd);
                            if (resulta != "OK")
                            {
                                MessageBox.Show(resulta, "Error en actualización de tabla usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Application.Exit();
                                return;
                            }
                            // a.code,a.names,a.passw,a.central,a.count,b.Descrizione,a.blocked
                            DataRow nrow = dtg.NewRow();
                            nrow["code"] = tx_codigo.Text;
                            nrow["passw"] = tx_pass.Text;
                            nrow["names"] = tx_usweb.Text;
                            nrow["central"] = "LIM";
                            nrow["count"] = tx_dat_cuenta.Text;
                            nrow["blocked"] = checkBox1.Checked;
                            nrow["Descrizione"] = Tx_ctaDes.Text;
                            dtg.Rows.Add(nrow);
                        }
                        catch (MySqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error en ingresar usuario", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            iserror = "si";
                        }
                        conn.Close();
                    }
                    else
                    {
                        MessageBox.Show("No se estableció conexión con el servidor", "Atención - no se puede continuar");
                        Application.Exit();
                        return;
                    }
                }
            }
            if (modo == "EDITAR")
            {
                var mes = MessageBox.Show("Realmente desea MODIFICAR el usuario web?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (mes == DialogResult.Yes)
                {
                    string consulta = "update webuser set names=@p5txt2,passw=@p5txt3,blocked=@p5chk1,central=@p5com2,count=@p5com1," +
                        "userm=@use,fechm=now(),diriplan4=@dirl,diripwan4=@dirw,netbname=@nbna " +
                        "where code=@p5txt1";
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        MySqlCommand mycom = new MySqlCommand(consulta, conn);
                        mycom.Parameters.AddWithValue("@p5txt1", tx_codigo.Text);
                        mycom.Parameters.AddWithValue("@p5txt2", tx_usweb.Text);
                        mycom.Parameters.AddWithValue("@p5txt3", tx_pass.Text);
                        mycom.Parameters.AddWithValue("@p5com2", "LIM");
                        mycom.Parameters.AddWithValue("@p5com1", tx_dat_cuenta.Text);
                        mycom.Parameters.AddWithValue("@p5chk1", (checkBox1.Checked == true) ? 1 : 0);
                        mycom.Parameters.AddWithValue("@verap", verapp);
                        mycom.Parameters.AddWithValue("@use", asd);
                        mycom.Parameters.AddWithValue("@dirl", lib.iplan());
                        mycom.Parameters.AddWithValue("@dirw", Conti3.Program.vg_ipwan);
                        mycom.Parameters.AddWithValue("@nbna", Environment.MachineName);
                        try
                        {
                            mycom.ExecuteNonQuery();
                            string resulta = lib.ult_mov(nomform, nomtab, asd);
                            if (resulta != "OK")                                        // actualizamos la tabla usuarios
                            {
                                MessageBox.Show(resulta, "Error en actualización de usuarios web", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Application.Exit();
                                return;
                            }
                            // // a.code,a.names,a.passw,a.central,a.count,b.Descrizione,a.blocked
                            if (tx_rind.Text.Trim() != "")
                            {
                                dtg.Rows[int.Parse(tx_rind.Text)]["names"] = tx_usweb.Text;
                                dtg.Rows[int.Parse(tx_rind.Text)]["passw"] = tx_pass.Text;
                                dtg.Rows[int.Parse(tx_rind.Text)]["count"] = tx_dat_cuenta.Text;
                                dtg.Rows[int.Parse(tx_rind.Text)]["Descrizione"] = Tx_ctaDes.Text;    // cmb_destin.Text;
                                dtg.Rows[int.Parse(tx_rind.Text)]["blocked"] = (checkBox1.Checked == true) ? 1 : 0;
                            }
                            else
                            {
                                for (int i = dtg.Rows.Count - 1; i >= 0; i--)
                                {
                                    DataRow drX = dtg.Rows[i];
                                    if (drX["code"].ToString() == tx_codigo.Text.ToString())
                                    {
                                        dtg.Rows[i]["names"] = tx_usweb.Text;
                                        dtg.Rows[i]["passw"] = tx_pass.Text;
                                        dtg.Rows[i]["count"] = tx_dat_cuenta.Text;
                                        dtg.Rows[i]["Descrizione"] = Tx_ctaDes.Text;  // cmb_destin.Text;
                                        dtg.Rows[i]["blocked"] = (checkBox1.Checked == true) ? 1 : 0;
                                    }
                                }
                            }
                            dtg.AcceptChanges();    //
                        }
                        catch (MySqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error de Editar usuario web", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            iserror = "si";
                        }
                        conn.Close();
                    }
                    else
                    {
                        MessageBox.Show("No se estableció conexión con el servidor", "Atención - no se puede continuar");
                        Application.Exit();
                        return;
                    }
                }
            }
            if (iserror == "no")
            {
                // debe limpiar los campos y actualizar la grilla
                tabControl1.SelectedTab = tabuser;
                limpia_combos(tabuser);
                limpiapag(tabuser);
                limpia_otros(tabuser);
                tx_codigo.Focus();
                //dataload();
            }
        }
        #endregion boton_form;

        #region leaves and activates
        private void tx_idr_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "NUEVO" && tx_idr.Text != "")
            {
                jalaoc("tx_idr");               // jalamos los datos del registro
            }
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (tx_codigo.Text.Trim() != "" && Tx_modo.Text != "NUEVO")
            {
                int cta = 0;
                foreach(DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString() == tx_codigo.Text.Trim())
                    {
                        cta += 1;
                        tx_rind.Text = row.Cells["code"].RowIndex.ToString();
                        jalaoc("tx_idr");
                    }
                }
                if (cta == 0)
                {
                    tx_codigo.Text = "";
                }
            }
            if (tx_codigo.Text.Trim() != "" && Tx_modo.Text == "NUEVO")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString().Trim() == tx_codigo.Text.Trim())
                    {
                        MessageBox.Show("Esta repitiendo el código","Error",MessageBoxButtons.OK);
                        tx_codigo.Clear();
                        return;
                    }
                }
            }
        }
        private void textBox2_Leave(object sender, EventArgs e)
        {
            // 
        }
        private void tabgrilla_Enter(object sender, EventArgs e)
        {
            if (Tx_modo.Text.Trim() != "")
            {
                pintaFilaAnul();
            }
        }
        private void Tx_ctaDes_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDITAR")
            {
                xxx();
            }
        }
        private void Tx_ctaDes_KeyPress(object sender, KeyPressEventArgs e)
        {
           if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDITAR")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    xxx();
                }
            }
        }
        private void Tx_ctaDes_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        }
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        #endregion leaves;

        #region botones_de_comando_y_permisos  
        public void toolboton(MySqlConnection conn)
        {
            DataTable mdtb = new DataTable();
            const string consbot = "select * from permisos where formulario=@nomform and usuario=@use";
            //MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            //conn.Open();
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
                //finally { conn.Close(); }
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
                if (Convert.ToString(row["btn1"]) == "S")
                {
                    this.Bt_add.Visible = true;
                }
                else { this.Bt_add.Visible = false; }
                if (Convert.ToString(row["btn2"]) == "S")
                {
                    this.Bt_edit.Visible = true;
                }
                else { this.Bt_edit.Visible = false; }
                if (Convert.ToString(row["btn3"]) == "S")
                {
                    this.Bt_anul.Visible = true;
                }
                else { this.Bt_anul.Visible = false; }
                if (Convert.ToString(row["btn4"]) == "S")
                {
                    this.Bt_ver.Visible = true;
                }
                else { this.Bt_ver.Visible = false; }
                if (Convert.ToString(row["btn5"]) == "S")
                {
                    //    this.Bt_print.Visible = true;
                }
                //else { this.Bt_print.Visible = false; }
                if (Convert.ToString(row["btn6"]) == "S")
                {
                    this.Bt_close.Visible = true;
                }
                else { this.Bt_close.Visible = false; }
            }
            else
            {
                Bt_add.Visible = false;
                Bt_edit.Visible = false;
                Bt_anul.Visible = false;
                Bt_ver.Visible = false;
            }
        }
        #region botones
        private void Bt_add_Click(object sender, EventArgs e)
        {
            advancedDataGridView1.Enabled = true;
            tabControl1.SelectedTab = tabuser;
            escribe(this);
            this.Tx_modo.Text = "NUEVO";
            this.button1.Image = Conti3.Properties.Resources.save_negro40;
            limpiar(this);
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            escribePag(tabuser);
            //cmb_destin.SelectedIndex = -1;
            this.tx_codigo.Focus();
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            advancedDataGridView1.Enabled = true;
            tabControl1.SelectedTab = tabuser;
            escribe(this);
            Tx_modo.Text = "EDITAR";
            button1.Image = Conti3.Properties.Resources.save_negro40;
            limpiar(this);
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            escribePag(tabuser);
            jalaoc("tx_idr");
            //cmb_destin.SelectedIndex = -1;
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            sololee(this);
            this.Tx_modo.Text = "IMPRIMIR";
            this.button1.Image = Conti3.Properties.Resources.delete40;
            this.tx_codigo.Focus();
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            // 
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            advancedDataGridView1.Enabled = true;
            sololee(this);
            Tx_modo.Text = "VISUALIZAR";
            this.button1.Image = null;
            limpiar(this);
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            sololeePag(tabuser);
            tabControl1.SelectedTab = tabgrilla;
            advancedDataGridView1.ReadOnly = true;
            //cmb_destin.SelectedIndex = -1;
            advancedDataGridView1.Focus();
        }
        private void Bt_first_Click(object sender, EventArgs e)
        {
            limpiar(this);
            limpia_chk();
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            //--
            tx_idr.Text = lib.gofirts(nomtab);
            tx_idr_Leave(null, null);
        }
        private void Bt_back_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            limpiar(this);
            //--
            tx_idr.Text = lib.goback(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_next_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            limpiar(this);
            //--
            tx_idr.Text = lib.gonext(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_last_Click(object sender, EventArgs e)
        {
            limpiar(this);
            limpia_chk();
            limpiapag(tabuser);
            limpia_otros(tabuser);
            limpia_combos(tabuser);
            //--
            tx_idr.Text = lib.golast(nomtab);
            tx_idr_Leave(null, null);
        }
        #endregion botones;
        #endregion botones_de_comando_y_permisos

        #region comboboxes
        // no hay
        #endregion comboboxes

        #region advancedatagridview
        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)                  // filtro de las columnas
        {
            dtg.DefaultView.RowFilter = advancedDataGridView1.FilterString;
            pintaFilaAnul();
        }
        private void advancedDataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)            // almacena valor previo al ingresar a la celda
        {
            advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
        }
        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Tx_modo.Text != "" && Tx_modo.Text != "NUEVO")
            {
                if (e.ColumnIndex == 0)
                {
                    tabControl1.SelectedTab = tabuser;
                    limpiar(this);
                    limpiapag(tabuser);
                    limpia_otros(tabuser);
                    limpia_combos(tabuser);
                    tx_rind.Text = advancedDataGridView1.CurrentRow.Index.ToString();
                    jalaoc("tx_idr");
                }
            }
        }
        private void advancedDataGridView1_SortStringChanged(object sender, EventArgs e)
        {
            dtg.DefaultView.Sort = advancedDataGridView1.SortString;
            pintaFilaAnul();
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
                        if (lib.validac(noeta[0], noeta[1], e.FormattedValue.ToString()) == true)
                        {
                            // llama a libreria con los datos para el update - tabla,id,campo,nuevo valor
                            lib.actuac(nomtab, campo, e.FormattedValue.ToString(),advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                        }
                        else
                        {
                            MessageBox.Show("El valor no es válido para la columna", "Atención - Corrija");
                            e.Cancel = true;
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
