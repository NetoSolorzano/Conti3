using MySql.Data.MySqlClient;
using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Presentation;

namespace Conti3
{
    public partial class Finan_reps1 : Form1
    {
        string nomForm = "Finan_reps1";
        Finan_Egres oFEgres = new Finan_Egres();
        //DataTable dt_ctaPer = new DataTable();      // cuentas personales
        DataTable dt_ctaOmg = new DataTable();      // cuentas omg
        DataTable dt_provee = new DataTable();      // proveedores
        DataTable dt_camion = new DataTable();      // camiones
        DataTable dt_s = new DataTable();           // cabecera reporte cuentas saldo inicial
        DataTable dt_d = new DataTable();           // detalle reporte de cuentas
        DataTable dt_dd = new DataTable();          // detalle para rep omg cuenta con categoria
        //List<string> lista_CON = new List<string>();    // cuentas personales 24/01/2025
        List<string> lista_CAM = new List<string>();    // categorias
        cajDestino Ocajd = new cajDestino();        // Objeto caja de destino
        catEgresos Ocateg = new catEgresos();       // Objeto categoria
        ccolores OColores = new ccolores();
        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";port=" + login.port + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data +
            ";ConnectionLifeTime=" + login.ctl + ";";

        string repMoxCaj = "";      // reporte de movimientos por caja
        string titMoxCaj = "";      // titulo del reporte movimientos x caja
        string repResPer = "";      // reporte resumen de cuentas personales
        string titResPer = "";      // titulo del reporte resumen de cuentas personales
        string repGenOmg = "";      // reporte general casa OMG
        string repGenOmgR = "";     // reporte general casa OMG resumen 
        string titGenOmgR = "";     // titulo reporte general casa OMG resumen
        string titGenOmg = "";      // titulo del reporte general casa OMG
        string repGasCam = "";      // reporte gastos camiones
        string titGasCam = "";      // titulo del reporte gastos camiones
        string rutExp = "";         // ruta para las exportaciones en excel
        public Finan_reps1()
        {
            InitializeComponent();
            cargaDatos();                                             // carga data de los combos
            jalainfo();
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                oFEgres.jalacolores(conn, OColores, nomForm);
                toolboton(conn);
            }
            oFEgres.colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);   // "#a8faf2", "#d9fbf8", "#d9fbf8"
        }
        private void Finan_reps1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
        }
        private void Tx_fecha1_Click(object sender, EventArgs e)
        {
            var mtb = (MaskedTextBox)sender;
            mtb.Select(0, 0);
            mtb.Focus();
        }
        private void Tx_fecha2_Click(object sender, EventArgs e)
        {
            var mtb = (MaskedTextBox)sender;
            mtb.Select(0, 0);
            mtb.Focus();
        }
        private void jalainfo()
        {
            nomForm = this.Name;
            foreach (DataRow row1 in Program.dt_enlaces.Rows)
            {
                if (row1.ItemArray[1].ToString() == "Main")
                {
                    if (row1.ItemArray[2].ToString() == "rutas")
                    {
                        if (row1.ItemArray[3].ToString() == "exporta") rutExp = row1.ItemArray[5].ToString();
                    }
                }
                if (row1.ItemArray[1].ToString() == nomForm)
                {
                    if (row1.ItemArray[2].ToString() == "reportes")
                    {
                        if (row1.ItemArray[3].ToString() == "repMoxCaj0") repMoxCaj = row1.ItemArray[5].ToString();
                        if (row1.ItemArray[3].ToString() == "titMoxCaj0") titMoxCaj = row1.ItemArray[5].ToString().Trim();
                        if (row1.ItemArray[3].ToString() == "repGenOmg0") repGenOmg = row1.ItemArray[5].ToString();
                        if (row1.ItemArray[3].ToString() == "titGenOmg0") titGenOmg = row1.ItemArray[5].ToString().Trim();
                        if (row1.ItemArray[3].ToString() == "repGenOmgR") repGenOmgR = row1.ItemArray[5].ToString();
                        if (row1.ItemArray[3].ToString() == "titGenOmgR") titGenOmgR = row1.ItemArray[5].ToString().Trim();
                        if (row1.ItemArray[3].ToString() == "repGasCam0") repGasCam = row1.ItemArray[5].ToString();
                        if (row1.ItemArray[3].ToString() == "titGasCam0") titGasCam = row1.ItemArray[5].ToString().Trim();
                        if (row1.ItemArray[3].ToString() == "repResPer0") repResPer = row1.ItemArray[5].ToString();
                        if (row1.ItemArray[3].ToString() == "titResPer0") titResPer = row1.ItemArray[5].ToString().Trim();
                    }
                }
            }
        }
        private string f_claudia(string f_normal)       // formato fecha de Claudia
        {
            string retorna = "";
            string[] aa = f_normal.Split('/');
            retorna = aa[0] + "." + aa[1] + "." + aa[2];
            return retorna;
        }
        private void marcaSelecGrilla(string modo)
        {
            for (int i = 0; i < advancedDataGridView1.Rows.Count - 1; i++)
            {
                DataGridViewRow row = advancedDataGridView1.Rows[i];
                if (modo == "marca")
                {
                    row.Cells[0].Value = true;
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#94f5ba");
                }
                if (modo == "desmar")
                {
                    row.Cells[0].Value = false;
                    row.DefaultCellStyle.BackColor = Color.White;
                }
                if (modo == "color")
                {
                    if (row.Cells[0].Value.ToString() == "True")
                    {
                        row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#94f5ba");
                    }
                }
            }
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

        #region combos y checks
        private void cargaDatos()
        {
            //cmb_categ.SelectedIndex = -1;
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error de conexión al servidor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Application.Exit();
                    return;
                }
                if (conn.State == ConnectionState.Open)
                {
                    // monedas
                    DataRow[] depar = Program.dt_definic.Select("idtabella='MON' and numero=1 AND idcodice<>'MON003'","idcodice desc");
                    cmb_moneda.DataSource = depar.CopyToDataTable();
                    cmb_moneda.DisplayMember = "descrizionerid";
                    cmb_moneda.ValueMember = "idcodice";

                    // casa o sede
                    cmb_sede.Items.Add("OMG");
                    cmb_sede.Items.Add("PERSONALES");
                    // cta personal                    
                    //depar = Program.dt_definic.Select("idtabella='CON' and numero=1");    //  and numero=1"
                    //foreach (DataRow row in depar)
                    //{
                    //    lista_CON.Add(row["descrizione"].ToString().Trim().ToUpper());
                    //}
                    //Tx_ctaDest.Values = lista_CON.ToArray();
                    //cargaCPer("Todos");
                    // cta omg
                    /*
                    dt_ctaOmg.Columns.Add("idcodice");
                    dt_ctaOmg.Columns.Add("descrizionerid");
                    dt_ctaOmg.Columns.Add("descrizione");
                    cargaCOmg("Todos"); */
                    // camiones
                    dt_camion.Columns.Add("idcodice");
                    dt_camion.Columns.Add("descrizionerid");
                    dt_camion.Columns.Add("descrizione");
                    // proveedor
                    using (MySqlCommand micon = new MySqlCommand("select trim(idanagrafica) as idanagrafica,trim(ragionesociale) as ragionesociale from anag_for order by ragionesociale", conn))
                    {
                        using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                        {
                            da.Fill(dt_provee);
                            cmb_prov.DataSource = dt_provee;
                            cmb_prov.DisplayMember = "ragionesociale";
                            cmb_prov.ValueMember = "idanagrafica";
                        }
                    }
                    // categorias de egreso/ingreso
                    DataRow[] cate = Program.dt_definic.Select("idtabella='CAM' and numero=1", "descrizionerid ASC");
                    foreach (DataRow row in cate)
                    {
                        lista_CAM.Add(row["descrizione"].ToString().Trim().ToUpper());
                    }
                    Tx_catEgre.Values = lista_CAM.ToArray();
                    //cmb_categ.DataSource = cate.CopyToDataTable();
                    //cmb_categ.DisplayMember = "descrizionerid";
                    //cmb_categ.ValueMember = "idcodice";
                }
            }
        }
        private void cargaCPer(string quienes)  // Todos || Activos
        {
            DataRow[] depar;
            List<string> lista_CON = new List<string>();
            if (quienes == "Todos")
            {
                depar = Program.dt_definic.Select("idtabella='CON'");    //  and numero=1"
            }
            else
            {
                depar = Program.dt_definic.Select("idtabella='CON' and numero=1");
            }
            foreach (DataRow row in depar)
            {
                lista_CON.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_ctaDest.Values = lista_CON.ToArray();
        }
        private void cargaCOmg(string quienes)
        {
            DataRow[] _omgs = null;
            List<string> lista_DES = new List<string>();
            if (quienes == "Todos") _omgs = Program.dt_definic.Select("idtabella='DES'");    //  and numero=1"
            else _omgs = Program.dt_definic.Select("idtabella='DES' and numero=1");
            foreach (DataRow row in _omgs)
            {
                lista_DES.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_ctaDest.Values = lista_DES.ToArray();
        }
        private void cargaCCam(string quienes)
        {
            DataRow[] depar = null;
            if (quienes != "Todos") depar = Program.dt_definic.Select("idtabella='DES'");
            else depar = Program.dt_definic.Select("idtabella='DES' and numero=1");
            List<string> lista_ = new List<string>();
            foreach (DataRow row in depar)
            {
                lista_.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_ctaDest.Values = lista_.ToArray();
        }
        private void cmb_sede_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_sede.SelectedIndex > -1)
            {
                if (cmb_sede.Text == "OMG")
                {
                    if (chk_desAct.CheckState == CheckState.Checked) cargaCOmg("Todos");
                    else cargaCOmg("Activos");
                    //cmb_destin.DataSource = dt_ctaOmg;
                }
                if (cmb_sede.Text == "PERSONALES")
                {
                    if (chk_desAct.CheckState == CheckState.Checked) cargaCPer("Todos");
                    else cargaCPer("Activos");
                    //cmb_destin.DataSource = dt_ctaPer;
                }
                //cmb_destin.DisplayMember = "descrizione";    // "descrizionerid"
                //cmb_destin.ValueMember = "idcodice";
                //
                if (rb_gasCam.Checked == true)
                {
                    if (chk_desAct.CheckState == CheckState.Checked) cargaCCam("Todos");
                    else cargaCCam("Activos");
                    //cmb_destin.DataSource = dt_camion;
                }
            }
        }
        private void chk_desAct_CheckStateChanged(object sender, EventArgs e)
        {
            Tx_ctaDest.Text = "";
            Ocajd.codigo = ""; Ocajd.nombre = ""; Ocajd.largo = "";
            if (chk_desAct.CheckState == CheckState.Checked)
            {
                if (cmb_sede.Text == "OMG") cargaCOmg("Todos");
                if (cmb_sede.Text == "PERSONALES") cargaCPer("Todos");
            }
            else
            {
                if (cmb_sede.Text == "OMG") cargaCOmg("Activos");
                if (cmb_sede.Text == "PERSONALES") cargaCPer("Activos");
            }
        }
        private void rb_tos_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_tos.Checked == true)
                marcaSelecGrilla("marca");
        }
        private void rb_nin_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_nin.Checked == true)
                marcaSelecGrilla("desmar");
        }
        private void cmb_moneda_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // 02/04/2025 si es resumen de cuentas omg se limpiara la grilla porque el reporte
            // se genera solo en soles o dolares ... debe volver a generarse el rep y se cambia la moneda
            if (Tx_modo.Text == "IMPRIMIR" && cmb_moneda.SelectedIndex > -1)
            {
                if (rb_globOmg.Checked == true && Tx_ctaDest.Text == "")
                {
                    advancedDataGridView1.DataSource = null;
                    advancedDataGridView1.Columns.Clear();
                    advancedDataGridView1.Rows.Clear();
                }
            }
        }

        #endregion

        #region Botones de comando
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
                    consulb.Parameters.AddWithValue("@nomform", nomForm);
                    consulb.Parameters.AddWithValue("@use", Program.vg_user);
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
                if (Convert.ToString(row["btn1"]) == "S")   // add
                {
                    this.Bt_add.Visible = true;
                }
                else { this.Bt_add.Visible = false; }
                if (Convert.ToString(row["btn2"]) == "S")   // edit
                {
                    this.Bt_edit.Visible = true;
                }
                else { this.Bt_edit.Visible = false; }
                if (Convert.ToString(row["btn3"]) == "S")   // anul
                {
                    this.Bt_anul.Visible = true;
                }
                else { this.Bt_anul.Visible = false; }
                if (Convert.ToString(row["btn4"]) == "S")   // view
                {
                    this.Bt_ver.Visible = true;
                }
                else { this.Bt_ver.Visible = false; }
                if (Convert.ToString(row["btn5"]) == "S")   // print
                {
                    this.Bt_print.Visible = true;
                }
                else { this.Bt_print.Visible = false; }
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
                Bt_print.Visible = false;
            }
        }
        private void Bt_add_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "NUEVO";
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "EDICION";
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "BORRAR";
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "EXCEL";
            // segun la pestanha activa debe exportar
            string nombre = "";
            if (advancedDataGridView1.Rows.Count > 0)
            {
                string fec1 = Tx_fecha1.Text.Substring(6, 4) + "-" + Tx_fecha1.Text.Substring(3, 2) + "-" + Tx_fecha1.Text.Substring(0, 2);
                string fec2 = Tx_fecha2.Text.Substring(6, 4) + "-" + Tx_fecha2.Text.Substring(3, 2) + "-" + Tx_fecha2.Text.Substring(0, 2);
                if (rb_movCaja.Checked == true)
                {
                    nombre = "MovimientosCajaPersonal" + "_" + cmb_sede.Text + "_" + Tx_ctaDest.Text + "_" + fec1 + "_" + fec2;
                    nombre = nombre + "_" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "_" + DateTime.Now.TimeOfDay.Hours.ToString() + DateTime.Now.TimeOfDay.Minutes.ToString() + DateTime.Now.TimeOfDay.Seconds.ToString() + ".xlsx";
                    var aa = MessageBox.Show("Confirma que desea generar la hoja de calculo?",
                        "Archivo: " + nombre, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        var wb = new XLWorkbook();
                        DataTable dt = dt_d.Select("CHK=True").CopyToDataTable();
                        var ws = wb.Worksheets.Add(dt, "Conticassa");
                        //ws.Range("A1:A3").InsertRowsAbove(2);
                        //wb.Cells("A1").Value = "Titulo";
                        wb.SaveAs(rutExp + nombre);
                        MessageBox.Show("Archivo generado con exito!");
                        this.Close();
                    }
                }
                if (rb_globOmg.Checked == true)
                {
                    nombre = "RepsGlobalOMG" + "_" + cmb_sede.Text + "_" + Tx_ctaDest.Text + "_" + fec1 + "_" + fec2;
                    nombre = nombre + "_" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "_" + DateTime.Now.TimeOfDay.Hours.ToString() + DateTime.Now.TimeOfDay.Minutes.ToString() + DateTime.Now.TimeOfDay.Seconds.ToString() + ".xlsx";
                    var aa = MessageBox.Show("Confirma que desea generar la hoja de calculo?",
                        "Archivo: " + nombre, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        var wb = new XLWorkbook();
                        DataTable dt = dt_d.Select("CHK=True").CopyToDataTable();
                        var ws = wb.Worksheets.Add(dt, "Conticassa");
                        //ws.Range("A1:A3").InsertRowsAbove(2);
                        //wb.Cells("A1").Value = "Titulo";
                        wb.SaveAs(rutExp + nombre);
                        MessageBox.Show("Archivo generado con exito!");
                        this.Close();
                    }
                }
                if (rb_gasCam.Checked == true)
                {
                    nombre = "RepGastosCamiones" + "_" + cmb_sede.Text + "_" + Tx_ctaDest.Text + "_" + fec1 + "_" + fec2;
                    nombre = nombre + "_" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "_" + DateTime.Now.TimeOfDay.Hours.ToString() + DateTime.Now.TimeOfDay.Minutes.ToString() + DateTime.Now.TimeOfDay.Seconds.ToString() + ".xlsx";
                    var aa = MessageBox.Show("Confirma que desea generar la hoja de calculo?",
                        "Archivo: " + nombre, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        var wb = new XLWorkbook();
                        DataTable dt = dt_d.Select("CHK=True").CopyToDataTable();
                        var ws = wb.Worksheets.Add(dt, "Conticassa");
                        //ws.Range("A1:A3").InsertRowsAbove(2);
                        //wb.Cells("A1").Value = "Titulo";
                        wb.SaveAs(rutExp + nombre);
                        MessageBox.Show("Archivo generado con exito!");
                        this.Close();
                    }
                }
            }
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            panelGeneral1.Enabled = true;
            cmb_sede.Enabled = true;
            Tx_ctaDest.Enabled = true;
            Tx_catEgre.Enabled = true;
            cmb_prov.Enabled = false;
            Tx_fecha1.Enabled = true;
            Tx_fecha2.Enabled = true;
            cmb_moneda.Enabled = true;
            bt_genera.Enabled = true;
            bt_prev.Enabled = true;
            //
            Tx_modo.Text = "IMPRIMIR";
            Tx_fecha1.Text = "01/01/" + DateTime.Now.Year.ToString();
            Tx_fecha2.Text = DateTime.Now.Date.ToString();
        }
        private void Bt_ini_Click(object sender, EventArgs e)
        {
            // GO TOP
        }
        private void Bt_sig_Click(object sender, EventArgs e)
        {
            // SKIP 1
        }
        private void Bt_ret_Click(object sender, EventArgs e)
        {
            // SKIP -1
        }
        private void Bt_fin_Click(object sender, EventArgs e)
        {
            // GO BOTT
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        private void Finan_reps1_Load(object sender, EventArgs e)
        {
            panelGeneral1.Enabled = false;
            cmb_sede.Enabled = false;
            Tx_ctaDest.Enabled = false;
            Tx_catEgre.Enabled = false;
            cmb_prov.Enabled = false;
            Tx_fecha1.Enabled = false;
            Tx_fecha2.Enabled = false;
            cmb_moneda.Enabled = false;
            bt_genera.Enabled = false;
            bt_prev.Enabled = false;
        }

        #region radiobuttons
        private void rb_ctaPers_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_ctaPers.Checked == true)
            {
                cmb_sede.Enabled = false; cmb_sede.SelectedIndex = 1;
                Ocajd.codigo = "";
                Ocajd.nombre = "";
                Ocajd.largo = "";
                Tx_ctaDest.Enabled = false;
                Tx_ctaDest.Text = "";
                eti_categ.Visible = false;
                Tx_catEgre.Visible = false;
                Tx_catEgre.Enabled = false;
                Tx_catEgre.Text = "";
                Ocateg.codigo = ""; Ocateg.nombre = ""; Ocateg.largo = "";
                cmb_prov.Enabled = false; cmb_prov.SelectedIndex = -1;
                cmb_moneda.SelectedIndex = 0;
                advancedDataGridView1.DataSource = null;
                advancedDataGridView1.Columns.Clear();
                advancedDataGridView1.Rows.Clear();
            }
        }
        private void rb_movCaja_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_movCaja.Checked == true)
            {
                cmb_sede.Enabled = false;
                cmb_sede.SelectedIndex = 1; // PERSONALES
                Tx_ctaDest.Enabled = true;
                Tx_ctaDest.Text = "";
                Ocajd.codigo = ""; Ocajd.nombre = ""; Ocajd.largo = "";
                eti_categ.Visible = true;
                Tx_catEgre.Visible = true;
                Tx_catEgre.Enabled = true;
                Tx_catEgre.Text = "";
                Ocateg.codigo = ""; Ocateg.nombre = ""; Ocateg.largo = "";
                cmb_prov.Enabled = false; cmb_prov.SelectedIndex = -1;
                cmb_moneda.SelectedIndex = 0;
                advancedDataGridView1.DataSource = null;
                advancedDataGridView1.Columns.Clear();
                advancedDataGridView1.Rows.Clear();
            }
        }
        private void rb_globOmg_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_globOmg.Checked == true)
            {
                eti_categ.Visible = true;
                Tx_catEgre.Visible = true;
                Tx_catEgre.Enabled = true;
                Tx_catEgre.Text = "";
                if (chk_desAct.CheckState == CheckState.Checked) cargaCOmg("Todos");
                else cargaCOmg("Activos");
                cmb_sede.Enabled = false; cmb_sede.SelectedIndex = 0;
                Tx_ctaDest.Enabled = true;
                Tx_ctaDest.Text = "";
                Ocajd.codigo = ""; Ocajd.nombre = ""; Ocajd.largo = "";
                Ocateg.codigo = ""; Ocateg.nombre = ""; Ocateg.largo = "";
                cmb_prov.Enabled = true; cmb_prov.SelectedIndex = -1;
                cmb_moneda.SelectedIndex = 0;
                advancedDataGridView1.DataSource = null;
                advancedDataGridView1.Columns.Clear();
                advancedDataGridView1.Rows.Clear();
            }
        }
        private void rb_gasCam_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_gasCam.Checked == true)
            {
                eti_categ.Visible = false;
                Tx_catEgre.Visible = false;

                if (chk_desAct.CheckState == CheckState.Checked) cargaCCam("Todos");
                else cargaCCam("Activos");
                cmb_sede.Enabled = false; cmb_sede.SelectedIndex = 0;
                Tx_ctaDest.Enabled = true;
                Tx_ctaDest.Text = "";
                Ocajd.codigo = ""; Ocajd.nombre = ""; Ocajd.largo = "";
                Tx_catEgre.Enabled = false;
                Tx_catEgre.Text = "";
                cmb_prov.Enabled = true; cmb_prov.SelectedIndex = -1;
                cmb_moneda.SelectedIndex = 0;
                advancedDataGridView1.DataSource = null;
                advancedDataGridView1.Columns.Clear();
                advancedDataGridView1.Rows.Clear();
            }
        }
        #endregion

        #region validaciones
        private void Tx_fecha_Click(object sender, EventArgs e)
        {
            var mtb = (MaskedTextBox)sender;
            mtb.Select(0, 0);
            mtb.Focus();
        }
        private void Tx_fecha1_Validating(object sender, CancelEventArgs e)
        {
            if (Tx_fecha1.Text.Trim() != "/  /")
            {
                if (Tx_fecha1.Text.Trim().Length != 10)
                {
                    MessageBox.Show("Fecha incorrecta", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Tx_fecha1.Clear();
                    e.Cancel = true;
                }
            }
        }
        private void Tx_fecha2_Validating(object sender, CancelEventArgs e)
        {
            if (Tx_fecha2.Text.Trim() != "/  /")
            {
                if (Tx_fecha2.Text.Trim().Length != 10)
                {
                    MessageBox.Show("Fecha incorrecta", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Tx_fecha2.Clear();
                    e.Cancel = true;
                }
            }
        }
        private void cmb_moneda_Leave(object sender, EventArgs e)
        {
            bt_genera.Focus();
        }
        private void Tx_ctaDes_KeyPress(object sender, KeyPressEventArgs e)
        {
            /*if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
            {
                if (Tx_ctaDest.Text.Trim() != "" && Tx_ctaDest.Text.Length >= 3)
                {
                    //xxx();
                }
            }   */
        }
        private void Tx_ctaDes_Leave(object sender, EventArgs e)
        {
            if (Tx_ctaDest.Text.Trim() != "" && Tx_ctaDest.Text.Length >= 3)
            {
                string[] vuelto = null;
                if (rb_ctaPers.Checked == true || rb_movCaja.Checked == true)
                {
                    vuelto = oFEgres.ValiCtaCon(Tx_ctaDest.Text, "PER", "algo");
                }
                if (rb_gasCam.Checked == true)
                {
                    vuelto = oFEgres.ValiCtaCon(Tx_ctaDest.Text, "OMG", "algo");
                    Ocajd.codigo = vuelto[0];
                    Ocajd.nombre = vuelto[1];
                    Ocajd.largo = vuelto[2];
                }
                if (rb_globOmg.Checked == true)
                {
                    vuelto = oFEgres.ValiCtaCon(Tx_ctaDest.Text, "OMG", "algo");
                }
                if (rb_movCaja.Checked == true)   // vuelto.Length > 0 && vuelto[0] != ""
                {
                    if (Tx_ctaDest.Text.Length > 4 && vuelto[0] != "")
                    {
                        Ocajd.codigo = vuelto[0];
                        Ocajd.nombre = vuelto[1];
                        Ocajd.largo = vuelto[2];
                    }
                    else
                    {
                        Ocajd.codigo = "";
                        Ocajd.nombre = "";
                        Ocajd.largo = "";
                        if (Tx_ctaDest.Text.Length > 5)
                        {
                            Tx_ctaDest.Clear();
                            MessageBox.Show("No existe el nombre de la cuenta");
                        }
                    }
                }
                if (rb_globOmg.Checked == true)
                {
                    if (Tx_ctaDest.Text.Length > 3 && vuelto[0] != "")
                    {
                        Ocajd.codigo = vuelto[0];
                        Ocajd.nombre = vuelto[1];
                        Ocajd.largo = vuelto[2];
                    }
                    else
                    {
                        Ocajd.codigo = "";
                        Ocajd.nombre = "";
                        Ocajd.largo = "";
                        if (Tx_ctaDest.Text.Length > 3)
                        {
                            Tx_ctaDest.Clear();
                            MessageBox.Show("No existe el nombre de la cuenta");
                        }
                    }
                }
            }
            else
            {
                Ocajd.codigo = "";
                Ocajd.nombre = "";
                Ocajd.largo = "";
                Tx_ctaDest.Text = "";
            }
        }
        private void Tx_catEgre_Leave(object sender, EventArgs e)
        {
            if (Tx_catEgre.Text.Trim() == "")
            {
                Ocateg.codigo = "";
                Ocateg.largo = "";
                Ocateg.nombre = "";
            }
            else
            {
                DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catEgre.Text.Trim() + "'");
                if (nc.Length > 0)
                {
                    Ocateg.codigo = nc[0].ItemArray[1].ToString();
                    Ocateg.largo = nc[0].ItemArray[2].ToString();
                    Ocateg.nombre = nc[0].ItemArray[3].ToString();
                }
                else
                {
                    Tx_catEgre.Clear();
                    Ocateg.codigo = "";
                    Ocateg.largo = "";
                    Ocateg.nombre = "";
                    MessageBox.Show("No existe el nombre del egreso");
                }
            }
        }
        private void Tx_catEgre_KeyPress(object sender, KeyPressEventArgs e)
        {
            /*  if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
            {
                if (Tx_catEgre.Text.Trim() != "")
                {
                    DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catEgre.Text.Trim() + "'");
                    if (nc.Length > 0)
                    {
                        Ocateg.codigo = nc[0].ItemArray[1].ToString();
                        Ocateg.largo = nc[0].ItemArray[2].ToString();
                        Ocateg.nombre = nc[0].ItemArray[3].ToString();
                    }
                    else
                    {
                        Tx_catEgre.Clear();
                        MessageBox.Show("No existe el nombre del egreso");
                    }
                }
            }   */
        }
        private void TextBox_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        }
        #endregion

        private void bt_genera_Click(object sender, EventArgs e)
        {
            dt_s = new DataTable();           // cabecera reporte cuentas saldo inicial
            dt_d = new DataTable();           // detalle reporte de cuentas

            if (Tx_fecha1.Text.Trim() == "/  /")
            {
                MessageBox.Show("Debe ingresar una fecha inicial", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Tx_fecha1.Focus();
                return;
            }
            if (Tx_fecha2.Text.Trim() == "/  /")
            {
                MessageBox.Show("Debe ingresar una fecha final", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Tx_fecha2.Focus();
                return;
            }
            if (Tx_ctaDest.Text == "" && rb_movCaja.Checked == true)
            {
                MessageBox.Show("Debe seleccionar una cuenta", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Tx_ctaDest.Focus();
                return;
            }
            if (Ocajd.codigo == "" && rb_movCaja.Checked == true)
            {
                MessageBox.Show("Seleccione correctamente una cuenta", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Tx_ctaDest.Focus();
                return;
            }
            /*  if (Ocajd.codigo == "" && rb_globOmg.Checked == true)
            {
                MessageBox.Show("Seleccione correctamente una cuenta", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Tx_ctaDest.Focus();
                return;
            }   */
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (conn.State == ConnectionState.Open)
                {
                    dt_d.Clear();
                    dt_d.Columns.Clear();
                    advancedDataGridView1.DataSource = null;
                    advancedDataGridView1.Columns.Clear();

                    string va01 = "LIM";
                    string va02 = Tx_fecha1.Text.Substring(6, 4) + "-" + Tx_fecha1.Text.Substring(3, 2) + "-" + Tx_fecha1.Text.Substring(0, 2);
                    string va03 = Tx_fecha2.Text.Substring(6, 4) + "-" + Tx_fecha2.Text.Substring(3, 2) + "-" + Tx_fecha2.Text.Substring(0, 2);
                    string va02f = Tx_fecha1.Text.Substring(6, 4) + "-" + Tx_fecha1.Text.Substring(3, 2) + "-" + Tx_fecha1.Text.Substring(0, 2);
                    string va04 = (Ocateg.codigo != null) ? Ocateg.codigo : "";    // (cmb_categ.SelectedIndex > -1) ? cmb_categ.SelectedValue.ToString() : "";
                    string va05 = Ocajd.codigo;     // (cmb_destin.SelectedValue != null) ? cmb_destin.SelectedValue.ToString() : "";
                    if (rb_movCaja.Checked == true)
                    {
                        //string va06 = cmb_moneda.SelectedValue.ToString();
                        string consulta = "reps_saldoIniS";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@v_tabla", "cassaconti");
                            micon.Parameters.AddWithValue("@v_a01", va01);  // idconti = 'LIM'
                            micon.Parameters.AddWithValue("@v_a02", va02);  // fecha ini
                            micon.Parameters.AddWithValue("@v_a03", va03);  // fecha fin
                            micon.Parameters.AddWithValue("@v_a02f", va02f); // fecha 
                            micon.Parameters.AddWithValue("@v_a04", va04);  // categoria
                            micon.Parameters.AddWithValue("@v_a05", va05);  // cuenta
                            //micon.Parameters.AddWithValue("@v_a06", va06);  // moneda
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_s.Clear();
                                da.Fill(dt_s);
                            }
                        }
                        consulta = "reps_cuenta1";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@v_tabla", "cassaconti");
                            micon.Parameters.AddWithValue("@v_a01", va01);  // idconti = 'LIM'
                            micon.Parameters.AddWithValue("@v_a02", va02);  // fecha ini
                            micon.Parameters.AddWithValue("@v_a03", va03);  // fecha fin
                            micon.Parameters.AddWithValue("@v_a02f", va02f); // fecha 
                            micon.Parameters.AddWithValue("@v_a04", "");  // categoria
                            micon.Parameters.AddWithValue("@v_a05", va05);  // cuenta
                            //micon.Parameters.AddWithValue("@v_a06", va06);  // moneda
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                da.Fill(dt_d);
                                if (va04 != "")
                                {
                                    int asd = dt_d.Select("idcategoria='" + va04 + "'").Length;
                                    if (asd > 0)
                                    {
                                        dt_dd = dt_d.Select("idcategoria='" + va04 + "'").CopyToDataTable();
                                        DataColumn col = dt_dd.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                        col.SetOrdinal(0);
                                        advancedDataGridView1.DataSource = dt_dd;
                                    }
                                }
                                else
                                {
                                    DataColumn col = dt_d.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                    col.SetOrdinal(0);
                                    advancedDataGridView1.DataSource = dt_d;
                                }
                                grilla();
                            }
                        }
                    }     // Rep 1 - Movimiento por caja - Personal                    
                    if (rb_ctaPers.Checked == true)
                    {
                        string consulta = "reps_cuenta2";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@v_tabla", "cassaconti");
                            micon.Parameters.AddWithValue("@v_a01", (cmb_moneda.Text == "S/") ? "MON001" : "MON002");  // codigo moneda del reporte
                            micon.Parameters.AddWithValue("@v_a02", va02);  // fecha ini
                            micon.Parameters.AddWithValue("@v_a03", va03);  // fecha fin
                            micon.Parameters.AddWithValue("@v_a04", (chk_desAct.Checked == true) ? 0 : 1);  // estado de las cuentas 1=activo, vacío todos
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                da.Fill(dt_d);
                                DataColumn col = dt_d.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                col.SetOrdinal(0);
                                advancedDataGridView1.DataSource = dt_d;
                                grilla();
                            }
                        }
                    }     // Rep 2 - Resumen cuentas personales
                    if (rb_globOmg.Checked == true)
                    {
                        string va00 = (chk_desAct.CheckState == CheckState.Checked) ? "" : "1";
                        string va06 = (Ocateg.codigo != null) ? Ocateg.codigo : "";
                        string consulta = "reps_saldoIniS";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@v_tabla", "cassaomg");
                            micon.Parameters.AddWithValue("@v_a01", va01);  // idconti = 'LIM'
                            micon.Parameters.AddWithValue("@v_a02", va02);  // fecha ini
                            micon.Parameters.AddWithValue("@v_a03", va03);  // fecha fin
                            micon.Parameters.AddWithValue("@v_a02f", va02f); // fecha 
                            micon.Parameters.AddWithValue("@v_a04", va00);  // cuentas 1=activos, vacío=todos
                            micon.Parameters.AddWithValue("@v_a05", va05);  // cuenta
                            //micon.Parameters.AddWithValue("@v_a06", va06);  // moneda
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_s.Clear();
                                da.Fill(dt_s);
                            }
                        }
                        if (va05 != "") // reporte de una cuenta OMG
                        {
                            consulta = "reps_cuenta3c";
                            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                            {
                                micon.CommandType = CommandType.StoredProcedure;
                                micon.Parameters.AddWithValue("@v_tabla", "cassaomg");
                                micon.Parameters.AddWithValue("@v_a01", va01);      // idconti = 'LIM'
                                micon.Parameters.AddWithValue("@v_a02", va02);      // fecha ini
                                micon.Parameters.AddWithValue("@v_a03", va03);      // fecha fin
                                micon.Parameters.AddWithValue("@v_a02f", va02f);    // fecha 
                                micon.Parameters.AddWithValue("@v_a04", va00);      // cuentas 1=activos, vacío=todos
                                micon.Parameters.AddWithValue("@v_a05", va05);      // cuenta
                                micon.Parameters.AddWithValue("@v_a06", "");      // categoria
                                using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                                {
                                    da.Fill(dt_d);
                                    if (va06 != "")
                                    {
                                        if (dt_d.Rows.Count > 0)
                                        {
                                            if (dt_d.Select("idcategoria='" + va06 + "'").Length > 0)
                                            {
                                                dt_dd = dt_d.Select("idcategoria='" + va06 + "'").CopyToDataTable();
                                                DataColumn col = dt_dd.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                                col.SetOrdinal(0);
                                                advancedDataGridView1.DataSource = dt_dd;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        DataColumn col = dt_d.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                        col.SetOrdinal(0);
                                        advancedDataGridView1.DataSource = dt_d;
                                    }
                                    grilla();
                                }
                            }
                        }
                        else
                        {
                            // resumen de todas las cuentas OMG
                            if (va04 == "") // va04=categoria
                            {
                                consulta = "reps_cuenta3g";
                                using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                                {
                                    micon.CommandType = CommandType.StoredProcedure;
                                    micon.Parameters.AddWithValue("@v_tabla", "cassaomg");
                                    micon.Parameters.AddWithValue("@v_a01", va01);      // idconti = 'LIM'
                                    micon.Parameters.AddWithValue("@v_a02", va02);      // fecha ini
                                    micon.Parameters.AddWithValue("@v_a03", va03);      // fecha fin
                                    micon.Parameters.AddWithValue("@v_a04", (chk_desAct.CheckState == CheckState.Checked) ? 0 : 1);         // cuentas 0=todos, 1=activos
                                    using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                                    {
                                        da.Fill(dt_d);
                                        DataColumn col = dt_d.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                        col.SetOrdinal(0);
                                        advancedDataGridView1.DataSource = dt_d;
                                        grilla();
                                    }
                                }
                            }
                            else
                            {
                                consulta = "reps_cuenta3gc";    // "reps_cuenta3g";    
                                using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                                {
                                    micon.CommandType = CommandType.StoredProcedure;
                                    micon.Parameters.AddWithValue("@v_tabla", "cassaomg");
                                    micon.Parameters.AddWithValue("@v_a01", va01);      // idconti = 'LIM'
                                    micon.Parameters.AddWithValue("@v_a02", va02);      // fecha ini
                                    micon.Parameters.AddWithValue("@v_a03", va03);      // fecha fin
                                    //micon.Parameters.AddWithValue("@v_a04", (chk_desAct.CheckState == CheckState.Checked) ? 0 : 1);
                                    micon.Parameters.AddWithValue("@v_a00", (chk_desAct.CheckState == CheckState.Checked) ? 0 : 1);         // cuentas 0=todos, 1=activos
                                    micon.Parameters.AddWithValue("@v_a04", va04);
                                    using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                                    {
                                        da.Fill(dt_d);
                                        DataColumn col = dt_d.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                        col.SetOrdinal(0);
                                        advancedDataGridView1.DataSource = dt_d;
                                        grilla();
                                    }
                                }
                            }
                        }
                    }     // Rep 3 - Global - Casa OMG
                    if (rb_gasCam.Checked == true)
                    {
                        string consulta = "reps_saldoIni";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@v_tabla", "camion");
                            micon.Parameters.AddWithValue("@v_a01", va01);  // idconti = 'LIM'
                            micon.Parameters.AddWithValue("@v_a02", va02);  // fecha ini
                            micon.Parameters.AddWithValue("@v_a03", va03);  // fecha fin
                            micon.Parameters.AddWithValue("@v_a02f", va02f); // fecha 
                            micon.Parameters.AddWithValue("@v_a04", va04);  // categoria
                            micon.Parameters.AddWithValue("@v_a05", va05);  // cuenta
                            //micon.Parameters.AddWithValue("@v_a06", va06);  // moneda
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_s.Clear();
                                da.Fill(dt_s);
                            }
                        }
                        consulta = "reps_cuenta4";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@v_tabla", "camion");
                            micon.Parameters.AddWithValue("@v_a01", va01);  // idconti = 'LIM'
                            micon.Parameters.AddWithValue("@v_a02", va02);  // fecha ini
                            micon.Parameters.AddWithValue("@v_a03", va03);  // fecha fin
                            micon.Parameters.AddWithValue("@v_a02f", va02f); // fecha 
                            micon.Parameters.AddWithValue("@v_a04", va04);  // categoria
                            micon.Parameters.AddWithValue("@v_a05", va05);  // cuenta
                            //micon.Parameters.AddWithValue("@v_a06", va06);  // moneda
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                da.Fill(dt_d);
                                DataColumn col = dt_d.Columns.Add("CHK", System.Type.GetType("System.Boolean"));
                                col.SetOrdinal(0);
                                advancedDataGridView1.DataSource = dt_d;
                                grilla();
                            }
                        }
                    }      // Rep 4 - Gastos Camiones
                    Bt_ver.Visible = true;
                }
            }
        }

        #region ADVANCEDDATAGRID
        private void grilla()
        {
            for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
            {
                if (i == 0) advancedDataGridView1.Columns[i].ReadOnly = false;
                else advancedDataGridView1.Columns[i].ReadOnly = true;
                if (i > 99)
                {
                    advancedDataGridView1.Columns[i].Visible = false;
                }
                else
                {
                    advancedDataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    if (advancedDataGridView1.Rows[0].Cells[i].Value != null)
                    {
                        var isNumeric = decimal.TryParse(advancedDataGridView1.Rows[0].Cells[i].Value.ToString(), out _);
                        if (isNumeric == true) advancedDataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                }
            }
            int b = 0;
            for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
            {
                int a = advancedDataGridView1.Columns[i].Width;
                b += a;
                advancedDataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                advancedDataGridView1.Columns[i].Width = a;
            }
            for (int i = 0; i < advancedDataGridView1.Rows.Count - 1; i++)
            {
                advancedDataGridView1.Rows[i].Cells[0].Value = true;
            }
        }
        private void advancedDataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "CHK")
            {
                if (advancedDataGridView1.CurrentRow.Cells[0].Value.ToString() == "True")    // bool.Parse(advancedDataGridView1.CurrentRow.Cells[0].Value.ToString()) == true
                {
                    advancedDataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#94f5ba");    // Color.YellowGreen;
                }
                else
                {
                    advancedDataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                }
            }
        }
        private void advancedDataGridView1_SortStringChanged(object sender, EventArgs e)
        {
            DataTable dtg = (DataTable)advancedDataGridView1.DataSource;
            dtg.DefaultView.Sort = advancedDataGridView1.SortString;
            marcaSelecGrilla("color");
        }
        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)                  // filtro de las columnas
        {
            DataTable dtg = (DataTable)advancedDataGridView1.DataSource;
            dtg.DefaultView.RowFilter = advancedDataGridView1.FilterString;
            marcaSelecGrilla("color");
        }

        #endregion

        #region reportes
        private DataSet1 GeneraReporte1(string CR)
        {
            DataSet1 set1 = new DataSet1();
            if (rb_movCaja.Checked == true)
            {
                DataSet1.repSaldoIni_Row cabecera = set1.repSaldoIni_.NewrepSaldoIni_Row();
                cabecera.id = "0";
                cabecera.nomb_reporte = CR;
                cabecera.titu_reporte = titMoxCaj; // TITULO del reporte
                cabecera.nomb_cliente = Program.vg_cliente; // nombre del cliente de la aplicacion
                cabecera.logotipo = "";     // ruta y nombre del archivo logo del cliente de la aplicacion
                cabecera.cassa = cmb_sede.Text;    // OMG o PER
                cabecera.cuenta = Ocajd.codigo + " - " + Ocajd.nombre; // cmb_destin.SelectedValue.ToString();   // cta personal
                cabecera.nombre_cta = Ocajd.largo;      // nombre largo de la cuenta
                cabecera.nomCateg = Tx_catEgre.Text;
                cabecera.fecha_hasta = Tx_fecha1.Text;      // fecha hasta donde debe calcular el saldo
                cabecera.fecha_ini = Tx_fecha1.Text;        // fecha de inicio del reporte
                cabecera.fecha_fin = Tx_fecha2.Text;        // fecha de finalizacion del reporte
                cabecera.idBanco = "LIM";   // tenemos que ver que es esto de LIM porque no es PER o OMG
                /*  switch (cmb_moneda.Text)
                {
                    case "S/":
                        cabecera.antesS = decimal.Parse(dt_s.Rows[0].ItemArray[2].ToString());
                        // cabecera.antes = decimal.Parse(dt_s.Rows[0].ItemArray[2].ToString()) * decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());
                        break;
                    case "US$":
                        cabecera.antesD = decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());   // / decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());
                        break;
                    case "EUR":
                        cabecera.antes = 0;
                        break;
                }   */
                cabecera.antesS = decimal.Parse(dt_s.Rows[0].ItemArray[2].ToString());
                cabecera.antesD = decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());
                cabecera.simbMon = cmb_moneda.Text;
                cabecera.tipCam_D = decimal.Parse(dt_s.Rows[0].ItemArray[4].ToString());  // ya no usamos esto ... 1;
                // calculo de totales ingresos y egresos del periodo sin considerar si esta 
                // seleccionado o no la fila de la grilla
                decimal totIngS = 0; decimal totEgreS = 0;
                decimal totIngD = 0; decimal totEgreD = 0;
                foreach (DataRow row in dt_d.Rows)
                {
                    if (decimal.Parse(row["SOL_INGR"].ToString()) > 0) totIngS = totIngS + decimal.Parse(row["SOL_INGR"].ToString());
                    if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0) totEgreS = totEgreS + decimal.Parse(row["SOL_EGRE"].ToString());
                    if (decimal.Parse(row["USD_INGR"].ToString()) > 0) totIngD = totIngD + decimal.Parse(row["USD_INGR"].ToString());
                    if (decimal.Parse(row["EUR_INGR"].ToString()) > 0) totIngD = totIngD + decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                    if (decimal.Parse(row["USD_EGRE"].ToString()) > 0) totEgreD = totEgreD + decimal.Parse(row["USD_EGRE"].ToString());
                    if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0) totEgreD = totEgreD + decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                    /* switch (cmb_moneda.Text)
                    {
                        case "S/":
                            if (decimal.Parse(row["SOL_INGR"].ToString()) > 0) totIngS = totIngS + decimal.Parse(row["SOL_INGR"].ToString());
                            else
                            {
                            }
                            if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0) totEgreS = totEgreS + decimal.Parse(row["SOL_EGRE"].ToString());
                            else
                            {
                            }
                            break;
                        case "US$":
                            if (decimal.Parse(row["USD_INGR"].ToString()) > 0) totIngD = totIngD + decimal.Parse(row["USD_INGR"].ToString());
                            else
                            {
                                if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                {
                                    //totIng = totIng + decimal.Parse(row["EUR_INGR"].ToString()) * decimal.Parse(row["T_C"].ToString()); HORROROOOOOOSO !!!
                                    totIngD = totIngD + decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                }
                            }
                            if (decimal.Parse(row["USD_EGRE"].ToString()) > 0) totEgreD = totEgreD + decimal.Parse(row["USD_EGRE"].ToString());
                            else
                            {
                                if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                {
                                    totEgreD = totEgreD + decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                }
                            }
                            break;
                        case "EUR":
                            //totEgre = 0;    // euros no esta habilitado para seleccionr
                            //totIng = 0;     // como moneda del reporte ... 18/10/2024
                            break;
                    }   */
                }
                cabecera.total_IngS = totIngS;
                cabecera.total_IngD = totIngD;
                cabecera.total_EgrS = totEgreS;
                cabecera.total_EgrD = totEgreD;
                set1.repSaldoIni_.AddrepSaldoIni_Row(cabecera);
                // detalle del reporte
                foreach (DataRow row in (Tx_catEgre.Text == "") ? dt_d.Rows : dt_dd.Rows)  // DataRow row in dt_d.Rows
                {
                    if (row["CHK"].ToString() == "True")
                    {
                        DataSet1.detalle1Row detalle = set1.detalle1.Newdetalle1Row();
                        detalle.id = "0";
                        detalle.movimiento = row["IDMOV"].ToString();      // id movimiento
                        detalle.fecha = f_claudia(row["FECHA"].ToString());                // fecha
                        detalle.descrip = row["DESCRIPCION"].ToString();        // descripcion
                        detalle.giroconto = row["CTA_GIRO"].ToString();             // cta destino del giro
                        detalle.cta_giro = row["IDM_GIRO"].ToString();          // Id del movimiento del giro destino
                        switch (cmb_moneda.Text)
                        {
                            case "S/":
                                if (decimal.Parse(row["SOL_INGR"].ToString()) > 0) detalle.Sol_Ingre = decimal.Parse(row["SOL_INGR"].ToString());      // 
                                else
                                {
                                    // este ELSE nunca se va a ejecutar porque todos los ingresos incluso en Eu y $ tienen su equivalente en soles
                                    if (decimal.Parse(row["USD_INGR"].ToString()) > 0)
                                    {
                                        detalle.Sol_Ingre = decimal.Parse(row["SOL_INGR"].ToString());
                                    }
                                    if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                    {
                                        detalle.Sol_Ingre = decimal.Parse(row["SOL_INGR"].ToString());
                                    }
                                }
                                if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0) detalle.Sol_Egre = decimal.Parse(row["SOL_EGRE"].ToString());      // 
                                else
                                {
                                    // este ELSE nunca se va a ejecutar porque todos los ingresos incluso en Eu y $ tienen su equivalente en soles
                                    if (decimal.Parse(row["USD_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Sol_Egre = decimal.Parse(row["SOL_EGRE"].ToString());
                                    }
                                    if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Sol_Egre = decimal.Parse(row["SOL_EGRE"].ToString());
                                    }
                                }
                                break;
                            case "US$":
                                if (decimal.Parse(row["USD_INGR"].ToString()) > 0) detalle.Usd_Ingre = decimal.Parse(row["USD_INGR"].ToString());
                                else
                                {
                                    /*  if (decimal.Parse(row["SOL_INGR"].ToString()) > 0)
                                    {
                                        detalle.Usd_Ingre = decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                    }   */
                                    if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                    {
                                        detalle.Usd_Ingre = decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                    }
                                }
                                if (decimal.Parse(row["USD_EGRE"].ToString()) > 0) detalle.Usd_Egre = decimal.Parse(row["USD_EGRE"].ToString());
                                else
                                {
                                    /*  if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Usd_Egre = decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                    }   */
                                    if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Usd_Egre = decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                    }
                                }
                                break;
                            case "EUR":
                                detalle.Eur_Ingre = 0;
                                break;
                        }
                        detalle.Eur_Egre = 0;
                        detalle.categoria = row["CATEGORIA"].ToString();        // nombre de la categoria
                        detalle.cuenta = row["CUENTA"].ToString();              // cuenta del reporte
                        detalle.solesIng = decimal.Parse(row["SOL_INGR"].ToString());
                        detalle.SolesEgr = decimal.Parse(row["SOL_EGRE"].ToString());
                        set1.detalle1.Adddetalle1Row(detalle);
                    }
                }
            }
            if (rb_globOmg.Checked == true && Tx_ctaDest.Text.Trim() != "")
            {
                DataSet1.repSaldoIni_Row cabecera = set1.repSaldoIni_.NewrepSaldoIni_Row();
                cabecera.id = "0";
                cabecera.nomb_reporte = CR;
                cabecera.titu_reporte = titGenOmg; // TITULO del reporte
                cabecera.nomb_cliente = Program.vg_cliente; // nombre del cliente de la aplicacion
                cabecera.logotipo = "";     // ruta y nombre del archivo logo del cliente de la aplicacion
                cabecera.cassa = cmb_sede.Text;    // OMG o PER
                cabecera.cuenta = Ocajd.codigo + " - " + Ocajd.nombre; // cmb_destin.SelectedValue.ToString();   // cta personal
                cabecera.nombre_cta = Ocajd.largo;      // nombre largo de la cuenta
                cabecera.fecha_hasta = Tx_fecha1.Text;      // fecha hasta donde debe calcular el saldo
                cabecera.fecha_ini = Tx_fecha1.Text;        // fecha de inicio del reporte
                cabecera.fecha_fin = Tx_fecha2.Text;        // fecha de finalizacion del reporte
                cabecera.idBanco = "LIM";   // tenemos que ver que es esto de LIM porque no es PER o OMG
                /*  switch (cmb_moneda.Text)
                {
                    case "S/":
                        cabecera.antesS = decimal.Parse(dt_s.Rows[0].ItemArray[2].ToString());  // * decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());
                        break;
                    case "US$":
                        cabecera.antesD = decimal.Parse(dt_s.Rows[0].ItemArray[2].ToString());
                        break;
                    case "EUR":
                        //cabecera.antes = 0;
                        break;
                }   */
                cabecera.antesS = decimal.Parse(dt_s.Rows[0].ItemArray[2].ToString());
                cabecera.antesD = decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());
                cabecera.simbMon = cmb_moneda.Text;
                cabecera.tipCam_D = 1;  // ya no se usa .. 01/04/2025 decimal.Parse(dt_s.Rows[0].ItemArray[3].ToString());
                // calculo de totales ingresos y egresos del periodo sin considerar si esta 
                // seleccionado o no la fila de la grilla
                decimal totIngS = 0; decimal totEgreS = 0;
                decimal totIngD = 0; decimal totEgreD = 0;
                foreach (DataRow row in dt_d.Rows)
                {
                    if (decimal.Parse(row["SOL_INGR"].ToString()) > 0) totIngS = totIngS + decimal.Parse(row["SOL_INGR"].ToString());
                    if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0) totEgreS = totEgreS + decimal.Parse(row["SOL_EGRE"].ToString());
                    if (decimal.Parse(row["USD_INGR"].ToString()) > 0) totIngD = totIngD + decimal.Parse(row["USD_INGR"].ToString());
                    if (decimal.Parse(row["USD_EGRE"].ToString()) > 0) totEgreD = totEgreD + decimal.Parse(row["USD_EGRE"].ToString());
                    if (decimal.Parse(row["EUR_INGR"].ToString()) > 0) totIngD = totIngD + decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                    if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0) totEgreD = totEgreD + decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                    /* switch (cmb_moneda.Text)
                    {
                        case "S/":
                            if (decimal.Parse(row["SOL_INGR"].ToString()) > 0) totIngS = totIngS + decimal.Parse(row["SOL_INGR"].ToString());
                            else
                            {
                                if (decimal.Parse(row["USD_INGR"].ToString()) > 0)
                                {
                                    //totIng = totIng + decimal.Parse(row["USD_INGR"].ToString()) * decimal.Parse(row["T_C"].ToString());
                                    //totIng = totIng + decimal.Parse(row["SOL_INGR"].ToString());
                                }
                                if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                {
                                    //totIng = totIng + decimal.Parse(row["EUR_INGR"].ToString()) * decimal.Parse(row["T_C"].ToString());
                                    //totIng = totIng + decimal.Parse(row["SOL_INGR"].ToString());
                                }
                            }
                            if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0) totEgreS = totEgreS + decimal.Parse(row["SOL_EGRE"].ToString());
                            else
                            {
                                if (decimal.Parse(row["USD_EGRE"].ToString()) > 0)
                                {
                                    //totEgre = totEgre + decimal.Parse(row["USD_EGRE"].ToString()) * decimal.Parse(row["T_C"].ToString());
                                    //totEgre = totEgre + decimal.Parse(row["SOL_EGRE"].ToString());
                                }
                                if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                {
                                    //totEgre = totEgre + decimal.Parse(row["EUR_EGRE"].ToString()) * decimal.Parse(row["T_C"].ToString());
                                    //totEgre = totEgre + decimal.Parse(row["SOL_EGRE"].ToString());
                                }
                            }
                            break;
                        case "US$":
                            if (decimal.Parse(row["USD_INGR"].ToString()) > 0) totIngD = totIngD + decimal.Parse(row["USD_INGR"].ToString());
                            else
                            {
                                if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                {
                                    //totIng = totIng + decimal.Parse(row["EUR_INGR"].ToString()) * decimal.Parse(row["T_C"].ToString());
                                    //totIng = totIng + decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                }
                            }
                            if (decimal.Parse(row["USD_EGRE"].ToString()) > 0) totEgreD = totEgreD + decimal.Parse(row["USD_EGRE"].ToString());
                            else
                            {
                                if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                {
                                    //totEgre = totEgre + decimal.Parse(row["EUR_INGR"].ToString()) * decimal.Parse(row["T_C"].ToString());
                                    totEgreD = totEgreD + decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                }
                            }
                            break;
                        case "EUR":
                            //totEgre = 0;    // euros no esta habilitado para seleccionr
                            //totIng = 0;     // como moneda del reporte ... 18/10/2024
                            break;
                    } */
                }
                cabecera.total_IngS = totIngS;
                cabecera.total_EgrS = totEgreS;
                cabecera.total_IngD = totIngD;
                cabecera.total_EgrD = totEgreD;
                set1.repSaldoIni_.AddrepSaldoIni_Row(cabecera);
                // detalle del reporte
                foreach (DataRow row in (Tx_catEgre.Text == "") ? dt_d.Rows : dt_dd.Rows)
                {
                    if (row["CHK"].ToString() == "True")
                    {
                        DataSet1.detalle1Row detalle = set1.detalle1.Newdetalle1Row();
                        detalle.id = "0";
                        detalle.movimiento = row["IDMOV"].ToString();      // id movimiento
                        detalle.fecha = f_claudia(row["FECHA"].ToString());                // fecha
                        detalle.descrip = row["DESCRIPCION"].ToString();        // descripcion
                        //detalle.giroconto = row["GIRO"].ToString();             // cta destino del giro
                        //detalle.cta_giro = row["IDM_GIRO"].ToString();          // Id del movimiento del giro destino  
                        switch (cmb_moneda.Text)
                        {
                            case "S/":
                                if (decimal.Parse(row["SOL_INGR"].ToString()) > 0) detalle.Sol_Ingre = decimal.Parse(row["SOL_INGR"].ToString());      // 
                                else
                                {
                                    // este ELSE nunca se va a ejecutar porque todos los ingresos incluso en Eu y $ tienen su equivalente en soles
                                    if (decimal.Parse(row["USD_INGR"].ToString()) > 0)
                                    {
                                        detalle.Sol_Ingre = decimal.Parse(row["SOL_INGR"].ToString());
                                    }
                                    if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                    {
                                        detalle.Sol_Ingre = decimal.Parse(row["SOL_INGR"].ToString());
                                    }
                                }
                                if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0) detalle.Sol_Egre = decimal.Parse(row["SOL_EGRE"].ToString());      // 
                                else
                                {
                                    // este ELSE nunca se va a ejecutar porque todos los ingresos incluso en Eu y $ tienen su equivalente en soles
                                    if (decimal.Parse(row["USD_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Sol_Egre = decimal.Parse(row["SOL_EGRE"].ToString());
                                    }
                                    if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Sol_Egre = decimal.Parse(row["SOL_EGRE"].ToString());
                                    }
                                }
                                break;
                            case "US$":
                                if (decimal.Parse(row["USD_INGR"].ToString()) > 0) detalle.Usd_Ingre = decimal.Parse(row["USD_INGR"].ToString());
                                else
                                {
                                    /*  if (decimal.Parse(row["SOL_INGR"].ToString()) > 0)
                                    {
                                        detalle.Usd_Ingre = decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["T_C"].ToString());
                                    }   */
                                    if (decimal.Parse(row["EUR_INGR"].ToString()) > 0)
                                    {
                                        detalle.Usd_Ingre = decimal.Parse(row["SOL_INGR"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                    }
                                }
                                if (decimal.Parse(row["USD_EGRE"].ToString()) > 0) detalle.Usd_Egre = decimal.Parse(row["USD_EGRE"].ToString());
                                else
                                {
                                    /*  if (decimal.Parse(row["SOL_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Usd_Egre = decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["T_C"].ToString());
                                    }   */
                                    if (decimal.Parse(row["EUR_EGRE"].ToString()) > 0)
                                    {
                                        detalle.Usd_Egre = decimal.Parse(row["SOL_EGRE"].ToString()) / decimal.Parse(row["TC_DOL"].ToString());
                                    }
                                }
                                break;
                            case "EUR":
                                detalle.Eur_Ingre = 0;
                                break;
                        }
                        detalle.Eur_Egre = 0;
                        detalle.categoria = row["CATEGORIA"].ToString();        // nombre de la categoria
                        detalle.cuenta = row["CUENTA"].ToString();              // cuenta del reporte
                        detalle.solesIng = decimal.Parse(row["SOL_INGR"].ToString());
                        detalle.SolesEgr = decimal.Parse(row["SOL_EGRE"].ToString());
                        detalle.nomCta = row["DESTINATARIO"].ToString();        // nombre de la cuenta
                        set1.detalle1.Adddetalle1Row(detalle);
                    }
                }
            }       // reporte de cuenta omg
            if (rb_gasCam.Checked == true)
            {
                DataSet1.CabGasCamRow cabeza = set1.CabGasCam.NewCabGasCamRow();
                cabeza.destino = Ocajd.largo;
                cabeza.fecha_ini = Tx_fecha1.Text;
                cabeza.fecha_fin = Tx_fecha2.Text;
                cabeza.id = "0";
                cabeza.id_camion = "---";    // no recibimos este parametro en el formulario
                cabeza.nomb_reporte = repGasCam;
                cabeza.titu_rep = titGasCam;
                cabeza.tip_gasto = "---";    // no recibimos este parametro en el formulario
                set1.CabGasCam.AddCabGasCamRow(cabeza);
                //
                foreach (DataRow row in dt_d.Rows)
                {
                    if (row["CHK"].ToString() == "True")
                    {
                        DataSet1.DetGasCamRow detalle = set1.DetGasCam.NewDetGasCamRow();
                        detalle.id = "0";
                        string FC = row["FECHA"].ToString().Substring(6, 4).Substring(2, 2);
                        string FCc = row["FECHA"].ToString().Substring(0, 6) + FC;
                        detalle.fecha = f_claudia(FCc); // f_claudia(row["FECHA"].ToString())
                        detalle.movimiento = row["IDMOV"].ToString();
                        detalle.id_camion = row["ID_CAMION"].ToString();
                        detalle.cuenta = row["CUENTA"].ToString();
                        detalle.placa = row["ASIGNADO"].ToString();
                        detalle.destino = row["DESTINO"].ToString();
                        detalle.descrip = row["DESCRIPCION"].ToString().Trim();
                        detalle.combust = decimal.Parse(row["COMBUST"].ToString());
                        detalle.honorar = decimal.Parse(row["HONORAR"].ToString());
                        detalle.impuest = decimal.Parse(row["IMPTOS"].ToString());
                        detalle.repuest = decimal.Parse(row["RPTOS"].ToString());
                        detalle.varios = decimal.Parse(row["VARIOS"].ToString());
                        detalle.viaticos = decimal.Parse(row["VIATICOS"].ToString());
                        detalle.total = decimal.Parse(row["TOTAL"].ToString());
                        detalle.tipcamD = decimal.Parse(row["T_C"].ToString());
                        set1.DetGasCam.AddDetGasCamRow(detalle);
                    }
                }
            }
            if (rb_ctaPers.Checked == true)
            {
                foreach (DataRow row in dt_d.Rows)
                {
                    if (row["CHK"].ToString() == "True")
                    {
                        DataSet1.ResCtasPersRow detalle = set1.ResCtasPers.NewResCtasPersRow();
                        detalle.id = "0";
                        detalle.nomb_reporte = repResPer;
                        detalle.titu_rep = titResPer;
                        detalle.fecIni = Tx_fecha1.Text;
                        detalle.fecter = Tx_fecha2.Text;
                        detalle.codCaja = row["codigo"].ToString();
                        detalle.nomCaja = row["NOMBRE_LARGO"].ToString();
                        switch (cmb_moneda.Text)
                        {
                            case "S/":
                                detalle.sigMon = cmb_moneda.Text;
                                detalle.codMon = cmb_moneda.SelectedValue.ToString();
                                break;
                            case "US$":
                                detalle.sigMon = cmb_moneda.Text;
                                detalle.codMon = cmb_moneda.SelectedValue.ToString();
                                break;
                            case "EUR":
                                detalle.sigMon = "";
                                detalle.codMon = "";
                                break;
                        }
                        detalle.anterior = decimal.Parse(row["antes"].ToString());
                        detalle.ingresos = decimal.Parse(row["ingreso"].ToString());
                        detalle.salidas = decimal.Parse(row["salida"].ToString());
                        detalle.saldo = decimal.Parse(row["saldo"].ToString());
                        set1.ResCtasPers.AddResCtasPersRow(detalle);
                    }
                }
            }
            if (rb_globOmg.Checked == true && Tx_ctaDest.Text.Trim() == "")
            {
                foreach (DataRow row in dt_d.Rows)
                {
                    if (row["CHK"].ToString() == "True")
                    {
                        DataSet1.ResCtasPersRow detalle = set1.ResCtasPers.NewResCtasPersRow();
                        detalle.id = "0";
                        detalle.nomb_reporte = CR; 
                        detalle.titu_rep = titGenOmgR;
                        detalle.fecIni = Tx_fecha1.Text;
                        detalle.fecter = Tx_fecha2.Text;
                        detalle.codCaja = row["codigo"].ToString();
                        detalle.nomCaja = row["NOMBRE_LARGO"].ToString();
                        detalle.nomCateg = Tx_catEgre.Text; // row[].ToString();
                        switch (cmb_moneda.Text)
                        {
                            case "S/":
                                detalle.sigMon = cmb_moneda.Text;
                                detalle.codMon = cmb_moneda.SelectedValue.ToString();
                                break;
                            case "US$":
                                detalle.sigMon = cmb_moneda.Text;
                                detalle.codMon = cmb_moneda.SelectedValue.ToString();
                                break;
                            case "EUR":
                                detalle.sigMon = "";
                                detalle.codMon = "";
                                break;
                        }
                        detalle.anterior = decimal.Parse(row["antes"].ToString());
                        detalle.ingresos = decimal.Parse(row["ingreso"].ToString());
                        detalle.salidas = decimal.Parse(row["salida"].ToString());
                        detalle.saldo = decimal.Parse(row["saldo"].ToString());
                        set1.ResCtasPers.AddResCtasPersRow(detalle);
                    }
                }
            }       // resumen de cuentas omg
            return set1;
        }
        private void bt_prev_Click(object sender, EventArgs e)
        {
            DataSet1 set = null;
            if (advancedDataGridView1.Rows.Count > 0)
            {
                if (rb_movCaja.Checked == true)
                {
                    set = GeneraReporte1(repMoxCaj);    // "repCtaPers1.rpt"
                }
                if (rb_globOmg.Checked == true)
                {
                    if (Tx_ctaDest.Text.Trim() == "") set = GeneraReporte1(repGenOmgR);
                    else set = GeneraReporte1(repGenOmg);    // 
                }
                if (rb_gasCam.Checked == true)
                {
                    set = GeneraReporte1(repGasCam);
                }
                if (rb_ctaPers.Checked == true)
                {
                    set = GeneraReporte1(repResPer);
                }
                FrmVisual visual = new FrmVisual(set);
                visual.Show();
            }
        }
        #endregion

        private void Finan_reps1_Click(object sender, EventArgs e)
        {
            this.Activate();
            this.BringToFront();
        }
    }
}
