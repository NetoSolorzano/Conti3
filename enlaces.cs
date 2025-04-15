using System;
//using System.Configuration;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class enlaces : Form1
    {
        publicoConf lp = new publicoConf();
        libreria lib = new libreria();
        static string nomform = "enlaces"; // nombre del formulario
        static string nomtab = "enlaces";
        public int totfilgrid, cta;      // variables para impresion
        public string perAg = "";
        public string perMo = "";
        public string perAn = "";
        public string perIm = "";
        string v_tipAdm = "";
        string colpage = "#687086";   // color de los pageframes
        // string de conexion
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        DataTable dtg = new DataTable();

        public enlaces()
        {
            InitializeComponent();
        }
        private void enlaces_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N) Bt_add.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E) Bt_edit.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P) Bt_print.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.A) Bt_anul.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O) Bt_ver.PerformClick();
            //if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S) Bt_close.PerformClick();
        }
        private void enlaces_Load(object sender, EventArgs e)
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
            limpiar();
            sololee();
            dataload();
            grilla();
            this.KeyPreview = true;
            Bt_add.Enabled = false;
            tabControl1.SelectedTab = tabgrilla;
            advancedDataGridView1.Enabled = false;
            toolStrip1.BackColor = ColorTranslator.FromHtml(colpage);
            // color de boton Bt_graba
            button1.BackColor = ColorTranslator.FromHtml("#667d97");   //  "#656d77",   #f5510f", "#e76433"
            button1.Image = null;
        }
        private void init()
        {
            jalainfo();
        }
        private void grilla()                   // arma la grilla
        {
            // select id,formulario,campo,descrip,valor,param FROM coop2018.enlaces;
            Font tiplg = new Font("Arial",9, FontStyle.Bold);
            advancedDataGridView1.Font = tiplg;
            advancedDataGridView1.DefaultCellStyle.Font = tiplg;
            advancedDataGridView1.RowTemplate.Height = 15;
            advancedDataGridView1.DefaultCellStyle.BackColor = Color.CadetBlue;
            advancedDataGridView1.DataSource = dtg;
            // id 
            advancedDataGridView1.Columns[0].Visible = false;
            // formulario
            advancedDataGridView1.Columns[1].Visible = true;            // columna visible o no
            advancedDataGridView1.Columns[1].HeaderText = "FORMULARIO";    // titulo de la columna
            advancedDataGridView1.Columns[1].Width = 70;                // ancho
            advancedDataGridView1.Columns[1].ReadOnly = true;           // lectura o no
            advancedDataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // campo
            advancedDataGridView1.Columns[2].Visible = true;       
            advancedDataGridView1.Columns[2].HeaderText = "CAMPO";
            advancedDataGridView1.Columns[2].Width = 70;
            advancedDataGridView1.Columns[2].ReadOnly = true;          // las celdas de esta columna pueden cambiarse
            advancedDataGridView1.Columns[2].Tag = "validaNO";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // param
            advancedDataGridView1.Columns[3].Visible = true;
            advancedDataGridView1.Columns[3].HeaderText = "PARAMETRO";
            advancedDataGridView1.Columns[3].Width = 100;
            advancedDataGridView1.Columns[3].ReadOnly = true;
            advancedDataGridView1.Columns[3].Tag = "validaNO";          // las celdas de esta columna SI se validan
            advancedDataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // descrip
            advancedDataGridView1.Columns[4].Visible = true;
            advancedDataGridView1.Columns[4].HeaderText = "DESCRIPCION";
            advancedDataGridView1.Columns[4].Width = 130;
            advancedDataGridView1.Columns[4].ReadOnly = false;          // las celdas de esta columna pueden cambiarse
            advancedDataGridView1.Columns[4].Tag = "validaNO";          // las celdas de esta columna se validan
            advancedDataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            // valor
            advancedDataGridView1.Columns[5].Visible = true;       
            advancedDataGridView1.Columns[5].HeaderText = "VALOR";
            advancedDataGridView1.Columns[5].Width = 130;
            advancedDataGridView1.Columns[5].ReadOnly = false;
            advancedDataGridView1.Columns[5].Tag = "validaNO";          // las celdas de esta columna se NO se validan
            advancedDataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }
        private void jalainfo()                 // obtiene datos de imagenes
        {
            try
            {
                for (int t = 0; t < Program.dt_enlaces.Rows.Count; t++)
                {
                    DataRow row = Program.dt_enlaces.Rows[t];
                    /* if (row["campo"].ToString() == "imagenes")
                    {
                        if (row["param"].ToString() == "img_btN") img_btN = row["valor"].ToString().Trim();         // imagen del boton de accion NUEVO
                        if (row["param"].ToString() == "img_btE") img_btE = row["valor"].ToString().Trim();         // imagen del boton de accion EDITAR
                        if (row["param"].ToString() == "img_btA") img_btA = row["valor"].ToString().Trim();         // imagen del boton de accion ANULAR/BORRAR
                        if (row["param"].ToString() == "img_btQ") img_btq = row["valor"].ToString().Trim();         // imagen del boton de accion SALIR
                        //if (row["param"].ToString() == "img_btP") img_btP = row["valor"].ToString().Trim();         // imagen del boton de accion IMPRIMIR
                        // boton de vista preliminar .... esta por verse su utlidad
                        if (row["param"].ToString() == "img_bti") img_bti = row["valor"].ToString().Trim();         // imagen del boton de accion IR AL INICIO
                        if (row["param"].ToString() == "img_bts") img_bts = row["valor"].ToString().Trim();         // imagen del boton de accion SIGUIENTE
                        if (row["param"].ToString() == "img_btr") img_btr = row["valor"].ToString().Trim();         // imagen del boton de accion RETROCEDE
                        if (row["param"].ToString() == "img_btf") img_btf = row["valor"].ToString().Trim();         // imagen del boton de accion IR AL FINAL
                        if (row["param"].ToString() == "img_gra") img_grab = row["valor"].ToString().Trim();         // imagen del boton grabar nuevo
                        if (row["param"].ToString() == "img_anu") img_anul = row["valor"].ToString().Trim();         // imagen del boton grabar anular
                    } */
                    if (row["campo"].ToString() == "tipoUser" && row["param"].ToString() == "admin") v_tipAdm = row["valor"].ToString().Trim();         // tipo usuario administrador
                }
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
            if (campo == "tx_idr" && tx_rind.Text != "")
            {
                textBox1.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[1].Value.ToString();  // formulario
                textBox2.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[2].Value.ToString();  // campo
                tx_param.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[3].Value.ToString();  // parametro                
                textBox3.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[4].Value.ToString();  // descrip
                tx_valor.Text = advancedDataGridView1.Rows[int.Parse(tx_rind.Text)].Cells[5].Value.ToString();  // valor
            }
            if (campo == "tx_corre")
            {

            }
        }
        public void dataload()                  // jala datos para los combos y la grilla
        {
            dtg.Clear();
            dtg = Program.dt_enlaces.Clone();  //.Fill(dtg);
            foreach (DataRow dr in Program.dt_enlaces.Rows)
            {
                if (true) dtg.Rows.Add(dr.ItemArray);
            }
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
        private void sololee()
        {
            lp.sololee(this);
        }
        private void escribe()
        {
            lp.escribe(this);
        }
        private void limpiar()
        {
            lp.limpiar(this);
        }
        private void limpia_chk()
        {
            lp.limpia_chk(this);
        }
        private void limpia_otros()
        {
            //
        }
        private void limpia_combos()
        {
            lp.limpia_cmb(this);
        }
        private void limpia_pag()
        {
            lp.limpiapag(tabreg);
        }
        #endregion limpiadores_modos;

        #region boton_form GRABA EDITA ANULA
        private void button1_Click(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "EDITAR")
            {
                return;
            }
            // validamos que los campos no esten vacíos
            if (this.textBox1.Text == "")
            {
                MessageBox.Show("Seleccione un Id Código", " Error! ");
                textBox1.Focus();
                return;
            }
            if (this.textBox2.Text == "")
            {
                //MessageBox.Show("Seleccione el código 2", " Error! ");
                //return;
            }
            // grabamos, actualizamos, etc
            string modo = this.Tx_modo.Text;
            string iserror = "no";
            string asd = Program.vg_user;
            string verapp = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
            if (modo == "NUEVO")
            {
                // 
            }
            if (modo == "EDITAR")
            {
                string consulta = "update enlaces set " +
                        "valor=@nval,descrip=@ndes,userm=@asd,fechm=@fec " +
                        "where id=@idc";
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    MySqlCommand mycom = new MySqlCommand(consulta, conn);
                    mycom.Parameters.AddWithValue("@idc", tx_idr.Text);
                    mycom.Parameters.AddWithValue("@nval", tx_valor.Text);
                    mycom.Parameters.AddWithValue("@ndes", textBox3.Text);
                    mycom.Parameters.AddWithValue("@asd", asd);
                    mycom.Parameters.AddWithValue("@fec", DateTime.Now);
                    try
                    {
                        mycom.ExecuteNonQuery();
                        // el log de movimientos
                        string resulta = lib.ult_mov(nomform, nomtab, asd);
                        if (resulta != "OK")                                        // actualizamos la tabla usuarios
                        {
                            MessageBox.Show(resulta, "Error en actualización de tabla usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Exit();
                            return;
                        }
                        // actualizamos la grilla
                        for(int i=0; i<dtg.Rows.Count; i++)
                        {
                            if (dtg.Rows[i][0].ToString() == tx_idr.Text)
                            {
                                dtg.Rows[i][4] = textBox3.Text;
                                dtg.Rows[i][5] = tx_valor.Text;
                            }
                        }
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error de Editar enlace",MessageBoxButtons.OK,MessageBoxIcon.Warning);
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
            if (modo == "ANULAR")       // opción para borrar
            { 
                // no se anulan, solo se habilitan o deshabilitan
            }
            if (iserror == "no")
            {
                // debe limpiar los campos y actualizar la grilla
                limpiar();
                limpia_pag();
                limpia_otros();
                textBox1.Focus();
            }
        }
        #endregion boton_form;

        #region leaves
        private void tx_idr_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "NUEVO" && tx_idr.Text != "")
            {
                jalaoc("tx_idr");               // jalamos los datos del registro
            }
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {

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
                    consulb.Parameters.AddWithValue("@use", Program.vg_user);
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
                if (Convert.ToString(row["btn5"]) == "S")
                {
                    //this.Bt_print.Visible = true;
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
                Bt_close.Visible = true;
            }
        }
        #region botones
        private void Bt_add_Click(object sender, EventArgs e)
        {
            // no 
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            if (true)    // Program.vg_tius == v_tipAdm
            {
                advancedDataGridView1.Enabled = true;
                advancedDataGridView1.ReadOnly = true;
                string rin = "";
                if (advancedDataGridView1.CurrentRow.Index > -1)
                {
                    rin = advancedDataGridView1.CurrentRow.Index.ToString();
                }
                tabControl1.SelectedTab = tabgrilla;
                escribe();
                Tx_modo.Text = "EDITAR";
                button1.Image = Conti3.Properties.Resources.save_negro40;
                limpiar();
                limpia_otros();
                limpia_combos();
                limpia_pag();
                tx_idr.Text = rin;
                jalaoc("tx_idr");
            }
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            // no
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            // no
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabgrilla;
            Tx_modo.Text = "VISUALIZAR";
            limpiar();
            limpia_otros();
            limpia_combos();
            limpia_pag();
            sololee();
            button1.Image = null;
            advancedDataGridView1.Enabled = true;
            advancedDataGridView1.ReadOnly = true;
        }
        private void Bt_first_Click(object sender, EventArgs e)
        {
            limpiar();
            limpia_chk();
            limpia_combos();
            //--
            //tx_idr.Text = lib.gofirts(nomtab);
            tx_idr_Leave(null, null);
        }
        private void Bt_back_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpia_combos();
            limpiar();
            //--
            //tx_idr.Text = lib.goback(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_next_Click(object sender, EventArgs e)
        {
            string aca = tx_idr.Text;
            limpia_chk();
            limpia_combos();
            limpiar();
            //--
            //tx_idr.Text = lib.gonext(nomtab, aca);
            tx_idr_Leave(null, null);
        }
        private void Bt_last_Click(object sender, EventArgs e)
        {
            limpiar();
            limpia_chk();
            limpia_combos();
            //--
            //tx_idr.Text = lib.golast(nomtab);
            tx_idr_Leave(null, null);
        }
        #endregion botones;
        //permisos para habilitar los botones de comando
        #endregion botones_de_comando  ;

        #region advancedatagridview
        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)                  // filtro de las columnas
        {
            dtg.DefaultView.RowFilter = advancedDataGridView1.FilterString;
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
                tx_rind.Text = advancedDataGridView1.CurrentRow.Index.ToString();
                tabControl1.SelectedTab = tabreg;
                limpiar();
                limpia_otros();
                limpia_combos();
                tx_idr.Text = idr;
                jalaoc("tx_idr");
            }
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
                        if (validac(noeta[0], noeta[1], e.FormattedValue.ToString()) == true)
                        {
                            // llama a libreria con los datos para el update - tabla,id,campo,nuevo valor
                            actuac(nomtab, campo, e.FormattedValue.ToString(),advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
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
                        actuac(nomtab, campo, e.FormattedValue.ToString(), advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                    }
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }
        #endregion

        public bool validac(string tabla, string campo, string valor)       // retorna true si el valor del campo existe en la tabla
        {
            bool retorna = false;
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string consulta = "select count(*) from " + tabla + " where " + campo + "='" + valor + "'";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    if (dr.GetInt16(0) == 1) retorna = true;
                    else retorna = false;
                }
                dr.Close();
                conn.Close();
            }
            else
            {
                MessageBox.Show("No se pudo conectar al servidor", "Error de red");
                Application.Exit();
                retorna = false;
            }
            return retorna;
        }

        public bool actuac(string tabla, string campo, string valor, string id)       // retorna true si actualizó
        {
            bool retorna = false;
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    string consulta = "update " + tabla + " set " + campo + "='" + valor + "' where id = '" + id + "'";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.ExecuteNonQuery();
                    retorna = true;
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "No se actualizó el campo");
                    retorna = false;
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("No se pudo conectar al servidor", "Error de red");
                Application.Exit();
                retorna = false;
            }
            return retorna;
        }

    }
}
