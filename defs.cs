using System;
using System.Collections.Generic;

//using System.Configuration;
using System.Data;
using System.Drawing;
//using System.Web.UI;
using System.Windows.Forms;
//using ADGV;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class defs : Form
    {
        static string nomform = "defs"; // nombre del formulario
        string asd = Conti3.Program.vg_user;   // usuario conectado al sistema
        string verapp = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
        string colback = "#c0ccea";   // color de fondo del form
        string colpage = "#687086";   // color de los pageframes
        string colgrid = "#a9acb6";   // color fondo sin grillas
        string colstrp = "#687eb6";   // color fuerte
        string colanul = "#fc658b";    // color de las filas anuladas de la grilla
        static string nomtab = "descrittive";
        public int totfilgrid, cta;      // variables para impresion
        public string perAg = "";
        public string perMo = "";
        public string perAn = "";
        public string perIm = "";
        string idsv = "'CON','DES','CAM','CMN'";    // identif.  mostrarse y trabajarse
        List<string> lista_ = new List<string>();   // lista según el idtabella
        libreria lib = new libreria();
        Finan_Egres OFegre = new Finan_Egres();
        // string de conexion
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        DataTable dtg = new DataTable();

        public defs()
        {
            InitializeComponent();
        }
        private void defs_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
        }
        private void defs_Load(object sender, EventArgs e)
        {
            ToolTip toolTipNombre = new ToolTip();           // Create the ToolTip and associate with the Form container.
            // Set up the delays for the ToolTip.
            toolTipNombre.AutoPopDelay = 5000;
            toolTipNombre.InitialDelay = 1000;
            toolTipNombre.ReshowDelay = 500;
            toolTipNombre.ShowAlways = true;                 // Force the ToolTip text to be displayed whether or not the form is active.
            toolTipNombre.SetToolTip(toolStrip1, nomform);   // toolStrip1 Set up the ToolTip text for the object
            init();
            toolboton();
            limpiar(tabreg);
            sololee(tabreg);
            dataload();
            grilla();
            KeyPreview = true;
            OFegre.colorea(this, colstrp, colpage, colback);    // pinta el mundo de colores!
            tabControl1.SelectedTab = tabgrilla;
            advancedDataGridView1.Enabled = false;
            toolStrip1.BackColor = ColorTranslator.FromHtml(colpage);
            // color de boton Bt_graba
            button1.BackColor = ColorTranslator.FromHtml("#667d97");   //  "#656d77",   #f5510f", "#e76433"
            button1.Image = null;
        }
        private void init()
        {
            //this.BackColor = ColorTranslator.FromHtml(colback);
            toolStrip1.BackColor = ColorTranslator.FromHtml(colstrp);
            //tabreg.BackColor = ColorTranslator.FromHtml(colpage);
            advancedDataGridView1.DefaultCellStyle.BackColor = Color.CadetBlue; // ColorTranslator.FromHtml(colgrid);

            jalainfo();
            Bt_add.Image = Properties.Resources.new_tab20; // Image.FromFile(img_btN);
            Bt_edit.Image = Properties.Resources.edit20;
            Bt_anul.Image = Properties.Resources.delete20;
            Bt_ver.Image = Properties.Resources.search_left20; // Image.FromFile(img_btV);
            Bt_close.Image = Properties.Resources.close20;
            Bt_ini.Image = Properties.Resources.arrow_in_left20;
            Bt_sig.Image = Properties.Resources.arrow_right20;
            Bt_ret.Image = Properties.Resources.arrow_left20;
            Bt_fin.Image = Properties.Resources.arrow_in_right20;

            textBox1.CharacterCasing = CharacterCasing.Upper;
            textBox1.MaxLength = 6;
            textBox2.MaxLength = 6;
            textBox3.MaxLength = 45;
            textBox5.MaxLength = 15;
            tx_det1.MaxLength = 90;
            tx_det2.MaxLength = 45;
            tx_det3.MaxLength = 45;
            tx_det4.MaxLength = 45;
            tx_det5.MaxLength = 6;
            tx_enla1.MaxLength = 6;
            tx_csunat.MaxLength = 6;
        }
        private void grilla()                   // arma la grilla
        {
            Font tiplg = new Font("Arial",7, FontStyle.Bold);
            advancedDataGridView1.Font = tiplg;
            advancedDataGridView1.DefaultCellStyle.Font = tiplg;
            advancedDataGridView1.RowTemplate.Height = 15;
            //advancedDataGridView1.DefaultCellStyle.BackColor = Color.MediumAquamarine;
            advancedDataGridView1.DataSource = dtg;
            // id 
            advancedDataGridView1.Columns[0].Visible = false;
            // idtabela
            advancedDataGridView1.Columns[1].Visible = true;            // columna visible o no
            advancedDataGridView1.Columns[1].HeaderText = "TABLA";    // titulo de la columna
            advancedDataGridView1.Columns[1].Width = 150;                // 70  ancho
            advancedDataGridView1.Columns[1].ReadOnly = true;           // lectura o no
            advancedDataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // idcodice
            advancedDataGridView1.Columns[2].Visible = true;       
            advancedDataGridView1.Columns[2].HeaderText = "CODIGO";
            advancedDataGridView1.Columns[2].Width = 100;
            advancedDataGridView1.Columns[2].ReadOnly = true;          // las celdas de esta columna pueden cambiarse
            advancedDataGridView1.Columns[2].Tag = "validaNO";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // CODIGO2
            advancedDataGridView1.Columns[3].Visible = false;        // true
            advancedDataGridView1.Columns[3].HeaderText = "CODIGO2";
            advancedDataGridView1.Columns[3].Width = 100;
            advancedDataGridView1.Columns[3].ReadOnly = false;          // las celdas de esta columna pueden cambiarse
            advancedDataGridView1.Columns[3].Tag = "validaNO";          // las celdas de esta columna se validan
            advancedDataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // nombre
            advancedDataGridView1.Columns[4].Visible = true;       
            advancedDataGridView1.Columns[4].HeaderText = "DESCRIPCION";
            advancedDataGridView1.Columns[4].Width = 300;
            advancedDataGridView1.Columns[4].ReadOnly = false;
            advancedDataGridView1.Columns[4].Tag = "validaNO";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // nombre corto
            advancedDataGridView1.Columns[5].Visible = true;       
            advancedDataGridView1.Columns[5].HeaderText = "DESC.CORTA";
            advancedDataGridView1.Columns[5].Width = 100;
            advancedDataGridView1.Columns[5].ReadOnly = false;
            advancedDataGridView1.Columns[5].Tag = "validano";          // las celdas de esta columna SI se validan
            advancedDataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // ACTIVO
            advancedDataGridView1.Columns[6].Visible = true;    
            advancedDataGridView1.Columns[6].HeaderText = "ACTIVO";
            advancedDataGridView1.Columns[6].Width = 40;
            advancedDataGridView1.Columns[6].ReadOnly = true;
            advancedDataGridView1.Columns[6].Tag = "validaSI";
            advancedDataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // invisibles
            advancedDataGridView1.Columns[7].Visible = false;            // detalle 1
            advancedDataGridView1.Columns[8].Visible = false;            // detalle 2
            advancedDataGridView1.Columns[9].Visible = false;            // detalle 3
            advancedDataGridView1.Columns[10].Visible = false;            // detalle 4
            advancedDataGridView1.Columns[11].Visible = false;            // detalle 5 / ubigeo
            advancedDataGridView1.Columns[12].Visible = false;            // marca1
            advancedDataGridView1.Columns[13].Visible = false;            // marca2
            advancedDataGridView1.Columns[14].Visible = false;            // marca3
            advancedDataGridView1.Columns[15].Visible = false;            // enlace1
            advancedDataGridView1.Columns[16].Visible = false;            // codsunat
            advancedDataGridView1.Columns[17].Visible = false;            // idtabella
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
                    if (row["campo"].ToString() == "tipoUser")
                    {
                        // if (row["param"].ToString() == "admin") v_tipAdm = row["valor"].ToString().Trim();         // tipo de usuario administrador SISTEMAS
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
            //textBox1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[].Value.ToString();  // 
            textBox4.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[17].Value.ToString();  // idtabella
            checkBox1.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[6].Value.ToString() == "1") ? true : false;
            cargaTipo(textBox4.Text);
            textBox1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[2].Value.ToString();  // idcodice
            textBox2.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[3].Value.ToString();  // codigo 2
            textBox3.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[4].Value.ToString();  // descrizione
            aTx_desc.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[4].Value.ToString();  // descrizione
            textBox5.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[5].Value.ToString();  // descrizionerid
            tx_det1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[7].Value.ToString();  // detalle 1
            tx_det2.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[8].Value.ToString();  // detalle 2
            tx_det3.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[9].Value.ToString();  // detalle 3
            tx_det4.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[10].Value.ToString();  // detalle 4
            tx_det5.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[11].Value.ToString();  // detalle 5 / ubigeo
            if (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[1].Value.ToString() == "TPA")
            {
                chk_marc1.Checked = false;
                tx_enla1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[12].Value.ToString();  // enlace1
            }
            else
            {
                chk_marc1.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[12].Value.ToString() == "1") ? true : false;
                tx_enla1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[15].Value.ToString();  // enlace1
            }
            chk_marc2.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[13].Value.ToString() == "1") ? true : false;
            chk_marc3.Checked = (advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[14].Value.ToString() == "1") ? true : false;
            tx_csunat.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[16].Value.ToString();  // codigo sunat
            comboBox1.SelectedValue = textBox4.Text;
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
            // DATOS DEL COMBOBOX1  idtabella
            comboBox1.Items.Clear();
            //ComboItem citem_tpu = new ComboItem();
            string contpu = "select deta4,idtabella from descrittive " +    // select ' ',idtabella from descrittive
                "where idtabella in (" + idsv + ") group by idtabella";
            MySqlCommand cmbtpu = new MySqlCommand(contpu, conn);
            DataTable dttpu = new DataTable();
            MySqlDataAdapter datpu = new MySqlDataAdapter(cmbtpu);
            datpu.Fill(dttpu);
            comboBox1.DataSource = dttpu;
            comboBox1.DisplayMember = "deta4";  // idtabella
            comboBox1.ValueMember = "idtabella";
            // datos de las deficiones
            /*  string datgri = "select id,idtabella,idcodice,codigo,descrizione,descrizionerid,numero," +
                "deta1,deta2,deta3,deta4,ubidir,marca1,marca2,marca3,enlace1,codsunat " +
                "from descrittive order by idtabella,idcodice"; */
            string datgri = "select id,deta4,idcodice,codigo,descrizione,descrizionerid,numero," +
                "deta1,deta2,deta3,deta4,ubidir,marca1,marca2,marca3,enlace1,codsunat,idtabella " +
                "from descrittive where idtabella in (" + idsv + ") order by idtabella,idcodice";
            MySqlCommand cdg = new MySqlCommand(datgri, conn);
            MySqlDataAdapter dag = new MySqlDataAdapter(cdg);
            dtg.Clear();
            dag.Fill(dtg);
            //
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
        private void mascara(object sender, EventArgs e)    // cambia etiquetas según la definicion
        {
            lb_idcodice.Text = "Id Código";
            lb_codigo.Text = "Código2";
            lb_descriz.Text = "Descripción";
            lb_descrizrid.Text = "Descrip. Corta";
            lb_det1.Text = "Det. 1 (90 digi)";
            lb_det2.Text = "Det. 2 (45 digi)";
            lb_det3.Text = "Det. 3 (45 digi)";
            lb_det4.Text = "Det. 4 (45 digi)";
            lb_ubigeo.Text = "Det. 5 (6 digi)";
            chk_marc1.Text = "Marca1";
            chk_marc2.Text = "Marca2";
            chk_marc3.Text = "Marca3";
            lb_enla1.Text = "Enlace 1";
            switch (textBox4.Text)
            {
                case "LOC":
                    lb_idcodice.Text = "Id Código";
                    lb_codigo.Text = "Código2";
                    lb_descriz.Text = "Descripción";
                    lb_descrizrid.Text = "Descrip. Corta";
                    lb_det1.Text = "Dirección";
                    lb_det2.Text = "Departamt";
                    lb_det3.Text = "Provincia";
                    lb_det4.Text = "Distrito";
                    lb_ubigeo.Text = "Ubigeo";
                    chk_marc1.Text = "Usa Pre Guías";
                    chk_marc2.Text = "Usa Num.GR Automat.";
                    chk_marc3.Text = "Usa Num.VALES Automat.";
                    lb_enla1.Text = "Zona destino";
                    break;
                case "TPA":
                    lb_idcodice.Text = "Id Código";
                    lb_codigo.Text = "Código2";
                    lb_descriz.Text = "Descripción";
                    lb_descrizrid.Text = "Descrip. Corta";
                    lb_enla1.Text = "Días plazo";
                    break;
            }
        }
        private void pintaFilaAnul()
        {
            for (int i = 0; i < advancedDataGridView1.Rows.Count - 1; i++)
            {
                if (advancedDataGridView1.Rows[i].Cells[6].Value.ToString() == "0")
                {
                    advancedDataGridView1.Rows[i].DefaultCellStyle.BackColor = ColorTranslator.FromHtml(colanul);
                }
                else { advancedDataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.CadetBlue; }
            }
        }
        private void cargaTipo(string tipo)         // jala el tipo de cuenta para el campo autocompletado
        {
            if (Tx_modo.Text != "NUEVO")
            {
                // tipo de cuenta
                DataRow[] depar = Program.dt_definic.Select("idtabella='" + tipo + "'");    // "idtabella='CON' and numero=1"
                lista_.Clear();
                lista_ = new List<string>();
                foreach (DataRow row in depar)
                {
                    lista_.Add(row["descrizione"].ToString().Trim().ToUpper());
                }
                aTx_desc.Values = lista_.ToArray();
            }
        }

        #region limpiadores_modos
        public void sololee(TabPage lfrm)
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
        public void escribe(TabPage efrm)
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
        public static void limpiar(TabPage ofrm)   // Form ofrm
        {
            foreach (Control oControls in ofrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
                if (oControls is AutoCompleteTextBox)
                {
                    oControls.Text = "";
                }
                if (oControls is CheckBox)
                {
                    CheckBox cb = (CheckBox)oControls;
                    cb.CheckState = CheckState.Unchecked;
                }
            }
        }
        public void limpiatab(TabPage tab)
        {
            tabControl1.SelectedTab = tab;
            foreach (Control oControls in tab.Controls)
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
            chk_marc1.Checked = false;
            chk_marc2.Checked = false;
            chk_marc3.Checked = false;
        }
        public void limpia_otros()
        {
            //
        }
        public void limpia_combos()
        {
            comboBox1.SelectedIndex = -1;
        }
        #endregion limpiadores_modos;

        #region boton_form GRABA EDITA ANULA
        private void button1_Click(object sender, EventArgs e)
        {
            // validamos que los campos no esten vacíos
            if (textBox1.Text.Trim() == "")
            {
                MessageBox.Show("Seleccione un Id Código", " Error! ");
                textBox1.Focus();
                return;
            }
            if (textBox2.Text.Trim() == "")
            {
                //MessageBox.Show("Seleccione el código 2", " Error! ");
                //textBox2.Focus();
                //return;
            }
            if (textBox5.Text.Trim() == "")
            {
                MessageBox.Show("Ingrese una descripción corta", " Dato obligatorio! ");
                textBox5.Focus();
                return;
            }
            if (this.comboBox1.Text == "")
            {
                MessageBox.Show("Seleccione el Id de Tabla", " Atención ");
                comboBox1.Focus();
                return;
            }
            // grabamos, actualizamos, etc
            string modo = this.Tx_modo.Text;
            string iserror = "no";
            if (modo == "NUEVO")
            {
                if (textBox3.Text.Trim() == "")
                {
                    MessageBox.Show("Ingrese una descripción", " Dato obligatorio! ");
                    textBox3.Focus();
                    return;
                }

                if (textBox4.Text.Trim() == "")
                {
                    MessageBox.Show("Confirme Id Tabla", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    comboBox1.Focus();
                    return;
                }
                if (textBox1.Text.Trim().Length <= 3 && (textBox4.Text == "CON" || textBox4.Text == "DES"))
                {
                    MessageBox.Show("El código debe tener mas de 3 digitos", "Verifique", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    textBox1.Focus();
                    return;
                }
                // valida que no este repitiendo el idcodice
                for (int i = 0; i < dtg.Rows.Count; i++)
                {
                    DataRow row = dtg.Rows[i];
                    if (row[17].ToString() == textBox4.Text && row[2].ToString() == textBox1.Text)
                    {
                        //id,idtabella,idcodice,codigo,descrizione,descrizionerid,numero
                        MessageBox.Show("Esta repitiendo el código", "Verifique", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        textBox1.Focus();
                        return;
                    }
                }
                var mes = MessageBox.Show("Realmente desea AGREGAR este nuevo registro?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (mes == DialogResult.Yes)
                {
                    string consulta = "insert into descrittive (" +
                        "idtabella,idcodice,codigo,descrizione,descrizionerid,numero,sede," +
                        "deta1,deta2,deta3,deta4,ubidir,marca1,marca2,marca3,enlace1,codsunat," +
                        "verApp,userc,fechc,diriplan4,diripwan4,netbname)" +
                        " values (" +
                        "@idt,@idc,@cod,@des,@der,@num,@sed," +
                        "@det1,@det2,@det3,@det4,@det5,@mar1,@mar2,@mar3,@enl1,@csun," +
                        "@veap,@asd,now(),@dipl,@dipw,@nbna)";
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        string idtu = "0";
                        MySqlCommand mycomand = new MySqlCommand(consulta, conn);
                        mycomand.Parameters.AddWithValue("@idt", textBox4.Text);
                        mycomand.Parameters.AddWithValue("@idc", textBox1.Text);
                        mycomand.Parameters.AddWithValue("@cod", textBox2.Text);
                        mycomand.Parameters.AddWithValue("@des", textBox3.Text);
                        mycomand.Parameters.AddWithValue("@der", textBox5.Text);
                        mycomand.Parameters.AddWithValue("@num", (checkBox1.Checked == true) ? "1" : "0");
                        mycomand.Parameters.AddWithValue("@sed", "LIM");    // sede por defecto
                        mycomand.Parameters.AddWithValue("@det1", tx_det1.Text);
                        mycomand.Parameters.AddWithValue("@det2", tx_det2.Text);
                        mycomand.Parameters.AddWithValue("@det3", tx_det3.Text);
                        mycomand.Parameters.AddWithValue("@det4", comboBox1.Text);    // tx_det4.Text
                        mycomand.Parameters.AddWithValue("@det5", tx_det5.Text);
                        if (textBox4.Text == "TPA")
                        {
                            mycomand.Parameters.AddWithValue("@mar1", tx_enla1.Text);
                            mycomand.Parameters.AddWithValue("@enl1", "");
                        }
                        else
                        {
                            mycomand.Parameters.AddWithValue("@mar1", (chk_marc1.Checked == true) ? "1" : "0");
                            mycomand.Parameters.AddWithValue("@enl1", tx_enla1.Text);
                        }
                        mycomand.Parameters.AddWithValue("@mar2", (chk_marc2.Checked == true) ? "1" : "0");
                        mycomand.Parameters.AddWithValue("@mar3", (chk_marc3.Checked == true) ? "1" : "0");
                        mycomand.Parameters.AddWithValue("@csun", tx_csunat.Text);
                        mycomand.Parameters.AddWithValue("@veap", verapp);
                        mycomand.Parameters.AddWithValue("@asd", asd);
                        mycomand.Parameters.AddWithValue("@dipl", lib.iplan());
                        mycomand.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                        mycomand.Parameters.AddWithValue("@nbna", Environment.MachineName);
                        try
                        {
                            mycomand.ExecuteNonQuery();
                            mycomand = new MySqlCommand("select last_insert_id()", conn);
                            MySqlDataReader dr0 = mycomand.ExecuteReader();
                            if (dr0.Read()) idtu = dr0.GetInt64(0).ToString();  // dr0.GetString(0);
                            dr0.Close();
                            string resulta = lib.ult_mov(nomform, nomtab, asd);
                            if (resulta != "OK")                                    // actualizamos la tabla usuarios
                            {
                                MessageBox.Show(resulta, "Error en actualización de tabla definiciones", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        catch (MySqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error en ingresar definición", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            iserror = "si";
                        }
                        conn.Close();
                        // insertamos en el datatable
                        DataRow dr = dtg.NewRow();
                        // id,idtabella,idcodice,codigo,descrizione,descrizionerid,numero,deta1,deta2,deta3,deta4,ubidir,marca1,marca2,marca3
                        dr[0] = idtu;
                        dr[1] = comboBox1.Text;   // 
                        dr[2] = textBox1.Text;
                        dr[3] = textBox2.Text;
                        dr[4] = textBox3.Text;
                        dr[5] = textBox5.Text;
                        dr[6] = (checkBox1.Checked == true) ? "1" : "0";
                        dr[7] = tx_det1.Text;
                        dr[8] = tx_det2.Text;
                        dr[9] = tx_det3.Text;
                        dr[10] = tx_det4.Text;
                        dr[11] = tx_det5.Text;
                        dr[12] = (textBox4.Text == "TPA") ? tx_enla1.Text : (chk_marc1.Checked == true) ? "1" : "0";
                        dr[13] = (chk_marc2.Checked == true) ? "1" : "0";
                        dr[14] = (chk_marc3.Checked == true) ? "1" : "0";
                        dr[15] = (textBox4.Text == "TPA") ? "" : tx_enla1.Text;
                        dr[16] = tx_csunat.Text;
                        dr[17] = textBox4.Text;
                        dtg.Rows.Add(dr);
                        // insertamos en dt_definic
                        DataRow drd = Program.dt_definic.NewRow();
                        drd[0] = textBox4.Text; 
                        drd[1] = textBox1.Text;
                        drd[2] = textBox3.Text;
                        drd[3] = textBox5.Text;
                        drd[4] = (checkBox1.CheckState == CheckState.Checked) ? "1" : "0";
                        Program.dt_definic.Rows.Add(drd);
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
                if (aTx_desc.Text.Trim() == "")
                {
                    MessageBox.Show("Ingrese una descripción", " Dato obligatorio! ");
                    textBox3.Focus();
                    return;
                }

                var mes = MessageBox.Show("Realmente desea EDITAR este registro?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (mes == DialogResult.Yes)
                {
                    string consulta = "update descrittive set " +
                        "descrizione=@des,descrizionerid=@der,numero=@num,codigo=@cod,sede=@sed," +
                        "deta1=@det1,deta2=@det2,deta3=@det3,deta4=@det4,ubidir=@det5," +
                        "marca1=@mar1,marca2=@mar2,marca3=@mar3,enlace1=@enl1,codsunat=@csun," +
                        "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                        "where idtabella=@idt and idcodice=@idc";    // id=@idc
                    MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        MySqlCommand mycom = new MySqlCommand(consulta, conn);
                        mycom.Parameters.AddWithValue("@cod", textBox2.Text);
                        mycom.Parameters.AddWithValue("@des", aTx_desc.Text);   // textBox3.Text
                        mycom.Parameters.AddWithValue("@der", textBox5.Text);
                        mycom.Parameters.AddWithValue("@num", (checkBox1.CheckState == CheckState.Checked) ? "1" : "0");  // checkBox1.Checked == true
                        mycom.Parameters.AddWithValue("@sed", "LIM");       // sede por defecto
                        mycom.Parameters.AddWithValue("@det1", tx_det1.Text);
                        mycom.Parameters.AddWithValue("@det2", tx_det2.Text);
                        mycom.Parameters.AddWithValue("@det3", tx_det3.Text);
                        mycom.Parameters.AddWithValue("@det4", comboBox1.Text); // tx_det4.Text
                        mycom.Parameters.AddWithValue("@det5", tx_det5.Text);
                        if (textBox4.Text == "TPA")
                        {
                            mycom.Parameters.AddWithValue("@mar1", tx_enla1.Text);
                            mycom.Parameters.AddWithValue("@enl1", "");
                        }
                        else
                        {
                            mycom.Parameters.AddWithValue("@mar1", (chk_marc1.Checked == true) ? "1" : "0");
                            mycom.Parameters.AddWithValue("@enl1", tx_enla1.Text);
                        }
                        mycom.Parameters.AddWithValue("@mar2", (chk_marc2.Checked == true) ? "1" : "0");
                        mycom.Parameters.AddWithValue("@mar3", (chk_marc3.Checked == true) ? "1" : "0");
                        mycom.Parameters.AddWithValue("@csun", tx_csunat.Text);
                        mycom.Parameters.AddWithValue("@veap", verapp);
                        mycom.Parameters.AddWithValue("@asd", asd);
                        mycom.Parameters.AddWithValue("@dipl", lib.iplan());
                        mycom.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                        mycom.Parameters.AddWithValue("@nbna", Environment.MachineName);
                        //mycom.Parameters.AddWithValue("@idc", tx_idr.Text);
                        mycom.Parameters.AddWithValue("@idt", textBox4.Text);
                        mycom.Parameters.AddWithValue("@idc", textBox1.Text);
                        try
                        {
                            mycom.ExecuteNonQuery();
                            string resulta = lib.ult_mov(nomform, nomtab, asd);
                            if (resulta != "OK")                                        // actualizamos la tabla usuarios
                            {
                                MessageBox.Show(resulta, "Error en actualización de tabla usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        catch (MySqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error de Editar definición", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            iserror = "si";
                        }
                        conn.Close();
                        // actualizamos dt_definic
                        DataRow[] dr = Program.dt_definic.Select("idtabella='" + textBox4.Text + "' and idcodice='" + textBox1.Text + "'");
                        if (dr != null)
                        {
                            dr[0]["descrizione"] = aTx_desc.Text;
                            dr[0]["descrizionerid"] = textBox5.Text;
                            dr[0]["numero"] = (checkBox1.CheckState == CheckState.Checked) ? "1" : "0";
                        }
                        // actualizamos el datatable
                        for (int i = 0; i < dtg.Rows.Count; i++)
                        {
                            DataRow row = dtg.Rows[i];  // row[0].ToString() == tx_idr.Text
                            if (row[2].ToString() == textBox1.Text && row[17].ToString() == textBox4.Text)
                            {
                                //id,idtabella,idcodice,codigo,descrizione,descrizionerid,numero
                                dtg.Rows[i][3] = textBox2.Text;
                                dtg.Rows[i][4] = aTx_desc.Text; // textBox3.Text
                                dtg.Rows[i][5] = textBox5.Text;
                                dtg.Rows[i][6] = (checkBox1.CheckState == CheckState.Checked) ? "1" : "0";
                                dtg.Rows[i][7] = tx_det1.Text;
                                dtg.Rows[i][8] = tx_det2.Text;
                                dtg.Rows[i][9] = tx_det3.Text;
                                dtg.Rows[i][10] = comboBox1.Text;     // tx_det4.Text
                                dtg.Rows[i][11] = tx_det5.Text;
                                dtg.Rows[i][12] = (textBox4.Text == "TPA") ? tx_enla1.Text : (chk_marc1.Checked == true) ? "1" : "0";
                                dtg.Rows[i][13] = (chk_marc2.Checked == true) ? "1" : "0";
                                dtg.Rows[i][14] = (chk_marc3.Checked == true) ? "1" : "0";
                                dtg.Rows[i][15] = (textBox4.Text == "TPA") ? "" : tx_enla1.Text;
                                dtg.Rows[i][16] = tx_csunat.Text;
                                dtg.Rows[i][17] = textBox4.Text;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se estableció conexión con el servidor", "Atención - no se puede continuar");
                        Application.Exit();
                        return;
                    }
                }
            }
            if (modo == "ANULAR")       // opción para borrar
            { 
                // no se anulan, solo se habilitan o deshabilitan
            }
            if (iserror == "no")
            {
                // debe limpiar los campos y actualizar la grilla
                limpiar(tabreg);
                limpiatab(tabreg);
                limpia_otros();
                limpia_chk();
                limpia_combos();
                //textBox1.Focus();
                comboBox1.Focus();
                pintaFilaAnul();
            }
        }
        #endregion boton_form;

        #region leaves and enters
        private void tx_idr_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "NUEVO" && tx_idr.Text != "")
            {
                jalaoc("tx_idr");               // jalamos los datos del registro
            }
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {
            //  validamos segun el modo
            if (textBox1.Text != "" && Tx_modo.Text=="NUEVO")
            {
                // buscar en la grilla
                DataRow[] ss = dtg.Select("idcodice='" + textBox1.Text.Trim() + "'");
                if (ss.Length > 0)
                {
                    MessageBox.Show("Código YA existe!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox1.Text = "";
                    return;
                }
            }
            if (textBox1.Text != "" && Tx_modo.Text != "NUEVO")
            {
                int contador = 0;
                //DataRow[] linea = dtg.Select("idcodice like '%" + textBox1.Text + "%' and idtabella='" + textBox4.Text + "'");
                DataRow[] linea = dtg.Select("idcodice='" + textBox1.Text.Trim() + "' and idtabella='" + textBox4.Text + "'");
                foreach (DataRow row in linea)
                {
                    contador = contador + 1;
                    textBox2.Text = row[3].ToString();
                    textBox3.Text = row[4].ToString();
                    aTx_desc.Text = row[4].ToString();
                    textBox5.Text = row[5].ToString();
                    checkBox1.Checked = (row[6].ToString() == "0") ? false : true;
                    tx_det1.Text = row[7].ToString();
                    tx_det2.Text = row[8].ToString();
                    tx_det3.Text = row[9].ToString();
                    tx_det4.Text = row[10].ToString();
                    tx_det5.Text = row[11].ToString();
                    chk_marc1.Checked = (row[12].ToString() == "0") ? false : true;
                    chk_marc2.Checked = (row[13].ToString() == "0") ? false : true;
                    chk_marc3.Checked = (row[14].ToString() == "0") ? false : true;
                    tx_enla1.Text = row[15].ToString();
                    tx_csunat.Text = row[16].ToString();
                    tx_idr.Text = row[0].ToString();     // id del registro en descrittive
                }
                if(contador == 0)
                {
                    MessageBox.Show("Código NO existe!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    aTx_desc.Text = "";
                    textBox5.Text = "";
                    checkBox1.Checked = false;
                    tx_det1.Text = "";
                    tx_det2.Text = "";
                    tx_det3.Text = "";
                    tx_det4.Text = "";
                    tx_det5.Text = "";
                    chk_marc1.Checked = false;
                    chk_marc2.Checked = false;
                    chk_marc3.Checked = false;
                    tx_enla1.Text = "";
                    tx_csunat.Text = "";
                    tx_idr.Text = "";
                    return;
                }
            }
        }
        private void tabgrilla_Enter(object sender, EventArgs e)
        {
            if (Tx_modo.Text.Trim() != "")
            {
                pintaFilaAnul();
            }
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void aTx_desc_TextChanged(object sender, EventArgs e)
        {
            if (aTx_desc.Text != "" && textBox4.Text == "")
            {
                aTx_desc.Text = "";
                MessageBox.Show("Debe seleccionar tipo de cuenta");
                comboBox1.Focus();
            }
        }
        private void aTx_desc_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
            {
                if (aTx_desc.Text.Trim() != "" && aTx_desc.Text.Length >= 3)
                {
                    if (Tx_modo.Text != "EDITAR")
                    {
                        if (textBox4.Text == "CAM") // tipos de egresos
                        {
                            DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            if (nc.Length > 0)
                            {
                                textBox1.Text = nc[0].ItemArray[1].ToString();   // idcodice
                                textBox5.Text = nc[0].ItemArray[3].ToString();   // descrip corta
                                checkBox1.Checked = (nc[0].ItemArray[4].ToString() == "1") ? true : false;
                            }
                            else
                            {
                                textBox1.Text = "";
                                textBox5.Text = "";
                                checkBox1.Checked = false;
                                MessageBox.Show("No existe el nombre del egreso");
                                aTx_desc.Text = "";
                            }
                        }
                        if (textBox4.Text == "CMN") // camion
                        {
                            DataRow[] row = Program.dt_definic.Select("idtabella='CMN' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            if (row.Length > 0)
                            {
                                textBox1.Text = row[0].ItemArray[1].ToString();   // idcodice
                                textBox5.Text = row[0].ItemArray[3].ToString();   // descrip corta
                                checkBox1.Checked = (row[0].ItemArray[4].ToString() == "1") ? true : false;
                            }
                            else
                            {
                                textBox1.Text = "";
                                textBox5.Text = "";
                                checkBox1.Checked = false;
                                MessageBox.Show("No existe el nombre del asignado");
                                aTx_desc.Text = "";
                            }
                        }
                        if (textBox4.Text == "CON" || textBox4.Text == "DES") // per, omg
                        {
                            DataRow[] vuelto = { };
                            if (textBox4.Text == "CON")
                            {
                                vuelto = Program.dt_definic.Select("idtabella='CON' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            }
                            if (textBox4.Text == "DES")
                            {
                                vuelto = Program.dt_definic.Select("idtabella='DES' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            }
                            if (vuelto.Length > 0)
                            {
                                textBox1.Text = vuelto[0].ItemArray[1].ToString();
                                textBox5.Text = vuelto[0].ItemArray[3].ToString();
                                checkBox1.Checked = (vuelto[0].ItemArray[4].ToString() == "1") ? true : false;
                            }
                            else
                            {
                                textBox1.Text = "";
                                textBox5.Text = "";
                                checkBox1.Checked = false;
                                MessageBox.Show("No existe el nombre de la cuenta");
                                aTx_desc.Text = "";
                            }
                        }
                    }
                    else
                    {
                        if (textBox4.Text == "CAM" && textBox1.Text.Trim() == "") // tipos de egresos
                        {
                            DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            if (nc.Length > 0)
                            {
                                textBox1.Text = nc[0].ItemArray[1].ToString();   // idcodice
                                textBox5.Text = nc[0].ItemArray[3].ToString();   // descrip corta
                                checkBox1.Checked = (nc[0].ItemArray[4].ToString() == "1") ? true : false;
                            }
                            else
                            {
                                textBox1.Text = "";
                                textBox5.Text = "";
                                checkBox1.Checked = false;
                                MessageBox.Show("No existe el nombre del egreso");
                                aTx_desc.Text = "";
                            }
                        }
                        if (textBox4.Text == "CMN" && textBox1.Text.Trim() == "") // camion
                        {
                            DataRow[] row = Program.dt_definic.Select("idtabella='CMN' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            if (row.Length > 0)
                            {
                                textBox1.Text = row[0].ItemArray[1].ToString();   // idcodice
                                textBox5.Text = row[0].ItemArray[3].ToString();   // descrip corta
                                checkBox1.Checked = (row[0].ItemArray[4].ToString() == "1") ? true : false;
                            }
                            else
                            {
                                textBox1.Text = "";
                                textBox5.Text = "";
                                checkBox1.Checked = false;
                                MessageBox.Show("No existe el nombre del asignado");
                                aTx_desc.Text = "";
                            }
                        }
                        if (textBox1.Text.Trim() == "" && (textBox4.Text == "CON" || textBox4.Text == "DES")) // per, omg
                        {
                            DataRow[] vuelto = { };
                            if (textBox4.Text == "CON")
                            {
                                vuelto = Program.dt_definic.Select("idtabella='CON' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            }
                            if (textBox4.Text == "DES")
                            {
                                vuelto = Program.dt_definic.Select("idtabella='DES' and descrizione='" + aTx_desc.Text.Trim() + "'");
                            }
                            if (vuelto.Length > 0)
                            {
                                textBox1.Text = vuelto[0].ItemArray[1].ToString();
                                textBox5.Text = vuelto[0].ItemArray[3].ToString();
                                checkBox1.Checked = (vuelto[0].ItemArray[4].ToString() == "1") ? true : false;
                            }
                            else
                            {
                                textBox1.Text = "";
                                textBox5.Text = "";
                                checkBox1.Checked = false;
                                MessageBox.Show("No existe el nombre de la cuenta");
                                aTx_desc.Text = "";
                            }
                        }
                    }
                }
            }
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
                //if (Convert.ToString(row["btn5"]) == "S")
                //{
                //    this.Bt_print.Visible = true;
                //}
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
                Bt_close.Visible = true;
            }
        }
        #region botones
        private void Bt_add_Click(object sender, EventArgs e)
        {
            if (true)    // Program.vg_tius == v_tipAdm
            {
                advancedDataGridView1.Enabled = true;
                advancedDataGridView1.ReadOnly = true;
                pintaFilaAnul();
                tabControl1.SelectedTab = tabreg;
                escribe(tabreg);
                Tx_modo.Text = "NUEVO";
                button1.Image = Properties.Resources.save_negro40; //Image.FromFile(img_grab);
                textBox1.Focus();
                limpiar(tabreg);
                limpiatab(tabreg);
                limpia_chk();
                limpia_otros();
                limpia_combos();
                //
                textBox3.Visible = true;
                aTx_desc.Visible = false;
                textBox3.TabIndex = 4;
                aTx_desc.TabIndex = 444;
                textBox3.Width = 393;
                textBox3.Left = 167;
                //
                button1.Enabled = true;
            }
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            if (true)    // Program.vg_tius == v_tipAdm
            {
                advancedDataGridView1.Enabled = true;
                advancedDataGridView1.ReadOnly = true;
                pintaFilaAnul();
                escribe(tabreg);
                Tx_modo.Text = "EDITAR";
                button1.Image = Properties.Resources.save_negro40; //Image.FromFile(img_grab);
                //var qa = tx_rind.Text;
                tabControl1.SelectedTab = tabgrilla;
                limpiar(tabreg);
                limpiatab(tabreg);
                //tx_rind.Text = qa;
                limpia_otros();
                limpia_combos();
                limpia_chk();
                //jalaoc("tx_idr");
                advancedDataGridView1.Focus();
                //
                textBox3.Visible = false;
                aTx_desc.Visible = true;
                aTx_desc.TabIndex = 4;
                textBox3.TabIndex = 444;
                aTx_desc.Width = 393;
                aTx_desc.Left = 167;
                //
                button1.Enabled = true;
            }
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            sololee(tabreg);
            advancedDataGridView1.Enabled = true;
            advancedDataGridView1.ReadOnly = true;
            pintaFilaAnul();
            escribe(tabreg);
            limpiar(tabreg);
            limpia_combos();
            Tx_modo.Text = "VISUALIZAR";
            button1.Image = null; //Image.FromFile(img_ver);
            //
            Bt_ini.Enabled = true;
            Bt_sig.Enabled = true;
            Bt_ret.Enabled = true;
            Bt_fin.Enabled = true;
            //
            textBox3.Visible = false;
            aTx_desc.Visible = true;
            aTx_desc.TabIndex = 4;
            textBox3.TabIndex = 444;
            aTx_desc.Width = 393;
            aTx_desc.Left = 167;
            //
            button1.Enabled = false;
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            /* sololee(this);
            Tx_modo.Text = "IMPRIMIR";
            button1.Image = Image.FromFile("print48");
            textBox1.Focus(); */
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            // no hay
        }
        private void Bt_first_Click(object sender, EventArgs e)
        {
            limpiar(tabreg);
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
            limpiar(tabreg);
            //--
            tx_idr.Text = lib.goback(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_next_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpia_combos();
            limpiar(tabreg);
            //--
            tx_idr.Text = lib.gonext(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_last_Click(object sender, EventArgs e)
        {
            limpiar(tabreg);
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
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)     // definición
        {
            // lo cambie por el changecommitted
        }
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > -1)
            {
                DataRow row = ((DataTable)comboBox1.DataSource).Rows[comboBox1.SelectedIndex];
                textBox4.Text = (string)row["idtabella"];
            }
            //limpia_combos();
            var xx = textBox4.Text;
            cargaTipo(xx);
            limpiatab(tabreg);
            limpia_chk();
            limpia_otros();
            textBox4.Text = xx;
        }
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
            if(e.ColumnIndex == 1)
            {
                //string codu = "";
                string idr = "";
                idr = advancedDataGridView1.CurrentRow.Cells[0].Value.ToString();                
                tabControl1.SelectedTab = tabreg;
                limpiar(tabreg);
                tx_rind.Text = advancedDataGridView1.CurrentRow.Index.ToString();
                limpia_otros();
                limpia_combos();
                jalaoc("tx_idr");
                tx_idr.Text = idr;
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

        private void defs_Click(object sender, EventArgs e)
        {
            this.Activate();
            this.BringToFront();
        }

    }
}
