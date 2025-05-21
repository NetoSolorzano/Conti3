using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using ADGV;
using MySql.Data.MySqlClient;

namespace Conti3
{
    public partial class Finan_Ingres : Form1
    {
        string nomform = "Finan_Ingres";
        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";port=" + login.port + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data +
            ";ConnectionLifeTime=" + login.ctl + ";";
        // datos de la grilla
        internal DataTable dt_grillaI = new DataTable();
        //
        publicoConf conf = new publicoConf();
        List<string> lista_CAM = new List<string>();                                    // categorias
        List<string> lista_DES = new List<string>();                                    // cuentas DES
        List<string> lista_CON = new List<string>();                                    // cuentas CON
        //
        catIngresos OcatIn = new catIngresos();                                     // Objeto categoría de egreso
        monedas Omone = new monedas();                                              // Objeto moneda
        tipcamDia tcDia = new tipcamDia();
        cajDestino Ocajd = new cajDestino();                                        // Objeto cada de destino - desde donde sale el dinero
        provees Oprove = new provees();                                             // Objeto proveedor
        montos Omonto = new montos();                                               // Objeto monto
        giroConto Ogiro = new giroConto();                                          // Objeto giroconto
        //
        Egresos oEgresos = new Egresos();
        Ingresos Oingreso = new Ingresos();
        Finan_Egres oFEgres = new Finan_Egres();
        ccolores OColores = new ccolores();
        string nomForm = "";
        int diasAtroya = 0;                                                         // dias atras hasta donde mostrará la grilla
        int limCols = 1;                                                            // limite de columnas que muestra la grilla
        string codDol = "MON002";
        string codEur = "MON003";
        string codSol = "MON001";
        string col1rafila = "";                                                     // color html de la 1ra fila en ingresos

        public Finan_Ingres()
        {
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            InitializeComponent();                  // inicializa los objetos graficos
            CargaINI(this);                         // colorea los objetos graficos
            CargaDatos();                           // jala datos de combos y demas
            chk_giroC_CheckedChanged(null, null);   // 
            sololee("T");                           // T=todos los campos, "" ó "C" campos comunes
            jalainfo();                             // jala variables de tabla enlace
            initCampos();                           // limita maximos de ancho en campos y mayusculas
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                oFEgres.jalacolores(conn, OColores, nomForm);
                toolboton(conn);
            }
            //oFEgres.colorea(this, "#89b174", "#badaa9", "#f2faed");   // pinta el mundo de colores
            oFEgres.colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);
            tx_descrip.BackColor = ColorTranslator.FromHtml(OColores.Fondo_suave);
            // color de boton Bt_graba
            //Bt_graba.BackColor = ColorTranslator.FromHtml("#667d97");
            Bt_graba.BackColor = ColorTranslator.FromHtml(OColores.Fondo_boton_graba);
            Bt_graba.Image = null;
            tx_diasA.Text = diasAtroya.ToString();
            bt_refresh.Image = Conti3.Properties.Resources.arrow_repeat15;
        }
        private void Finan_Ingres_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            string para1 = "";
            string para2 = "";
            string para3 = "";
            string para4 = "";
            if (keyData == Keys.F1 && (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION"))
            {
                if (Tx_ctaDest.Focused == true)
                {
                    para1 = (rb_omg.Checked == true) ? "omg" : "personal";
                    para2 = "cuenta";
                    para3 = "activos";    // todos | activos
                    ayuda2 ayu2 = new ayuda2(para1, para2, para3, para4);
                    var result = ayu2.ShowDialog();
                    if (result == DialogResult.Cancel)
                    {
                        if (!string.IsNullOrEmpty(ayu2.ReturnValue1))   // 0=codigo, 1=descripCorta, 2=descripLarga
                        {
                            Tx_ctaDest.Text = ayu2.ReturnValueA[2];     // ayu2.ReturnValueA[1]
                            eti_nomCaja.Text = ayu2.ReturnValueA[1];    // ayu2.ReturnValueA[2]
                            xxx();      // funcion que graba el objeto cta destino
                            bool v = tx_descrip.Focus();
                            if (v == true) Conti3.AutoClosingMessageBox.Show("Ingrese la descripción", "", 10);
                        }
                    }
                }
                if (tx_ctaGiro.Focused == true)
                {
                    para1 = (rb_omg.Checked == true) ? "omg" : "personal";
                    para2 = "cuenta";
                    para3 = "activos";    // todos | activos
                    ayuda2 ayu2 = new ayuda2(para1, para2, para3, para4);
                    var result = ayu2.ShowDialog();
                    if (result == DialogResult.Cancel)
                    {
                        if (!string.IsNullOrEmpty(ayu2.ReturnValue1))   // 0=codigo, 1=descripCorta, 2=descripLarga
                        {
                            tx_ctaGiro.Text = ayu2.ReturnValueA[2];     // [1]
                            eti_nomCtaGiro.Text = ayu2.ReturnValueA[1]; // [2]
                            tx_dat_giro.Text = ayu2.ReturnValueA[0];    // [0]
                        }
                    }
                }
                if (Tx_catIng.Focused == true)
                {
                    para1 = (rb_omg.Checked == true) ? "omg" : "personal";
                    para2 = "tEgresos";
                    para3 = "activos";    // todos | activos
                    ayuda2 ayu2 = new ayuda2(para1, para2, para3, para4);
                    var result = ayu2.ShowDialog();
                    if (result == DialogResult.Cancel)
                    {
                        if (!string.IsNullOrEmpty(ayu2.ReturnValue1))   // 0=codigo, 1=descripCorta, 2=descripLarga
                        {
                            Tx_catIng.Text = ayu2.ReturnValueA[2];  // A[1]
                            eti_nomCat.Text = ayu2.ReturnValueA[1]; // A[2]
                            OcatIn.codigo = ayu2.ReturnValueA[0];   // A[0]
                            OcatIn.nombre = ayu2.ReturnValueA[1];   // A[1]
                            OcatIn.largo = ayu2.ReturnValueA[2];    // A[2]
                        }
                    }
                }
                return true;    // indicate that you handled this keystroke
            }
            // Call the base class
            return base.ProcessCmdKey(ref msg, keyData);
        }    // F1 
        private void CargaDatos()
        {
            // categorias
            DataRow[] depar = Program.dt_definic.Select("idtabella='CAM' and numero=1");
            foreach (DataRow row in depar)
            {
                //lista_CAM.Add(row["descrizionerid"].ToString().Trim().ToUpper());
                lista_CAM.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_catIng.Values = lista_CAM.ToArray();
            // giro OMG
            depar = Program.dt_definic.Select("idtabella='DES' and numero=1");
            foreach (DataRow row in depar)
            {
                //lista_DES.Add(row["descrizionerid"].ToString().Trim().ToUpper());
                lista_DES.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            // giro PER
            depar = Program.dt_definic.Select("idtabella='CON' and numero=1");
            foreach (DataRow row in depar)
            {
                //lista_CON.Add(row["descrizionerid"].ToString().Trim().ToUpper());
                lista_CON.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            // monedas
            depar = Program.dt_definic.Select("idtabella='MON' and numero=1");
            cmb_mon.DataSource = depar.CopyToDataTable();
            cmb_mon.DisplayMember = "descrizionerid";
            cmb_mon.ValueMember = "idcodice";
        }
        private void jalainfo()
        {
            // 31/07/2024 .. variabilizamos los datos que vamos a necesitar
            nomForm = this.Name;
            DataRow[] row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='diasAtras'");
            diasAtroya = int.Parse(row[0]["valor"].ToString());
            row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='limCols'");
            limCols = int.Parse(row[0]["valor"].ToString());
            row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='col1rafila'");
            col1rafila = row[0]["valor"].ToString();              // color html de la 1ra fila en ingresos
        }
        private void jalaoc()
        {
            if (Tx_modo.Text != "NUEVO")
            {
                tx_idOper.Text = Oingreso.IdMovim;
                tx_anno.Text = Oingreso.AnnoOp;    // DateTime.Parse(Oingreso.FechOper).Year.ToString();
                selecFecha1.Value = DateTime.Parse(Oingreso.FechOper);
            }
            else
            {
                if (Oingreso.FechOper == "") selecFecha1.Value = DateTime.Now.Date;
                else selecFecha1.Value = DateTime.Parse(Oingreso.FechOper);
                if (Oingreso.AnnoOp == "") tx_anno.Text = DateTime.Now.Date.Year.ToString();
                else tx_anno.Text = Oingreso.AnnoOp;
            }
            Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
            Tx_catIng.Text = Oingreso.CatIngreso.largo;    // .nombre
            eti_nomCat.Text = Oingreso.CatIngreso.nombre;    // .largo
            cmb_mon.SelectedValue = Oingreso.Moneda.codigo;
            tx_monto.Text = Oingreso.Monto.monOrige.ToString("#0.00");
            tx_tipcam.Text = Oingreso.TipCamb.ToString("#0.000");
            Tx_ctaDest.Text = Oingreso.CajaDes.largo;  // .nombre
            eti_nomCaja.Text = Oingreso.CajaDes.nombre;  // .largo
            tx_descrip.Text = Oingreso.Descrip;
            tipCambio(null);
        }                                                   // muestra en el formulario los objetos de la clase Egresos
        private void initCampos()
        {
            tx_anno.MaxLength = 4;
            Tx_catIng.MaxLength = 50;   // 20
            Tx_catIng.CharacterCasing = CharacterCasing.Upper;
            Tx_ctaDest.MaxLength = 50;  // 20
            Tx_ctaDest.CharacterCasing = CharacterCasing.Upper;
            tx_ctaGiro.MaxLength = 50;  // 20
            tx_ctaGiro.CharacterCasing = CharacterCasing.Upper;
            tx_descrip.MaxLength = 93;  // 09/05/2025
            tx_descrip.Font = new Font(conf.nombreFont, conf.tamañoFont);
            tx_idOper.MaxLength = 15;
            tx_diasA.MaxLength = 3;
        }                                               // inicializa ancho de campos y upper case
        private void datsimil()
        {
            /*
            string[] ju = oFEgres.jala_ultimo(dt_grillaI, "INGRESO", ((rb_omg.Checked == true) ? rb_omg.Text.ToUpper().Substring(0,3) : rb_pers.Text.ToUpper().Substring(0, 3)), Tx_fecha.Text);
            //llenamos los objetos
            if (rb_omg.Checked == true && ju[3] != "")
            {
                Ocajd.largo = ju[8].ToString();
                Ocajd.nombre = ju[7].ToString();
                Ocajd.codigo = ju[6].ToString();
                OcatIn.largo = ju[3].ToString();
                OcatIn.nombre = ju[2].ToString();
                OcatIn.codigo = ju[1].ToString();
                Omone.nombre = ju[5].ToString();
                Omone.codigo = ju[4].ToString();
                Omone.siglas = ju[15].ToString();
                Omonto.tipCOri = decimal.Parse(ju[0]);
                Omonto.monOrige = decimal.Parse(ju[16]);
                Omonto.monSoles = decimal.Parse(ju[18]);
                Omonto.monDolar = decimal.Parse(ju[17]);
                //Omonto.monEuros = 
                Oprove = null;
                Ogiro.largo = "";
                Ogiro.idcod = "";
                Ogiro.ctades = "";
            }
            if (rb_pers.Checked == true && ju[3] != "")
            {
                Ocajd.largo = ju[8].ToString();
                Ocajd.nombre = ju[7].ToString();
                Ocajd.codigo = ju[6].ToString();
                OcatIn.largo = ju[3].ToString();
                OcatIn.nombre = ju[2].ToString();
                OcatIn.codigo = ju[1].ToString();
                Omone.nombre = ju[5].ToString();
                Omone.codigo = ju[4].ToString();
                Omone.siglas = ju[15].ToString();
                Omonto.tipCOri = decimal.Parse(ju[0]);
                Omonto.monOrige = decimal.Parse(ju[16]);
                Omonto.monSoles = decimal.Parse(ju[18]);
                Omonto.monDolar = decimal.Parse(ju[17]);
                //Omonto.monEuros = 
                Oprove = null;
                Ogiro.largo = "";
                Ogiro.idcod = "";
                Ogiro.ctades = "";
            }
            //
            Oingreso.creaIngreso(pan_p.Tag.ToString(), Tx_fecha.Text, OcatIn, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                Ocajd, ju[11].ToString(), "", Ogiro);
            jalaoc();   // pinta los datos en la pantalla
            */
        }
        private bool tipCambio(MySqlConnection condb)
        {
            bool retorna = false;
            MySqlConnection conn;
            if (condb == null)
            {
                conn = new MySqlConnection(DB_CONN_STR);
            }
            else
            {
                conn = condb;
            }
            try
            {
                conn.Open();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return retorna;
            }

            // buscamos tipo de cambio del día
            using (MySqlCommand micon = new MySqlCommand("select ifnull(Cambio1,0),ifnull(Cambio2,0) from cambi where date(datavaluta)=@fec", conn))  // dolares,euros
            {
                string fcv = selecFecha1.Value.ToString().Substring(6, 4) + "-" + selecFecha1.Value.ToString().Substring(3, 2) + "-" + selecFecha1.Value.ToString().Substring(0, 2);
                micon.Parameters.AddWithValue("@fec", fcv);
                using (MySqlDataReader dr = micon.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            if (Tx_modo.Text == "NUEVO")    // 20/05/2025
                            {
                                tx_tipcam.Text = Math.Round(dr.GetDecimal(0), 3).ToString();
                                Omonto.tipCDol = Math.Round(dr.GetDecimal(0), 3);
                                if (Omonto.codMOrige != null && Omonto.codMOrige != "")
                                {
                                    Omonto.tipCOri = (Omonto.codMOrige == codEur) ? Math.Round(dr.GetDecimal(1), 3) : (Omonto.codMOrige == codDol) ? Math.Round(dr.GetDecimal(0), 3) : Math.Round(dr.GetDecimal(0), 3);
                                }
                            }
                            tcDia.tcD = Omonto.tipCDol;
                            tcDia.tcE = Math.Round(dr.GetDecimal(1), 3);   // Omonto.tipCOri;

                            if (Omonto.tipCDol <= 0 || tcDia.tcE <= 0) // Omonto.tipCDol <= 0 || Omonto.tipCOri <= 0
                            {
                                MessageBox.Show("No existen tipos de cambio para la fecha actual" + Environment.NewLine +
                                    "Debe ingresarlos en este momento", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                retorna = false;
                            }
                            else { retorna = true; }
                        }
                    }
                    else
                    {
                        MessageBox.Show("No existen tipos de cambio para la fecha actual" + Environment.NewLine +
                            "Debe ingresarlos en este momento", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        retorna = false;
                    }
                }
            }
            return retorna;
        }

        #region Botones de comando
        public void toolboton(MySqlConnection conn)
        {
            DataTable mdtb = new DataTable();
            const string consbot = "select * from permisos where formulario=@nomform and usuario=@use";
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
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "NUEVO";
            if (tipCambio(null) == true)
            {
                if (rb_pers.Checked == false && rb_omg.Checked == false)
                {
                    rb_pers.Checked = true;
                    rb_pers_Click(null, null);
                }
                limpiaObj();
                limpiaTE();
                selecFecha1.Enabled = true;
                escribe("");
                tx_idOper.ReadOnly = true;
                selecFecha1.Value = DateTime.Now.Date; // DateTime.UtcNow.Date;
                Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
                tx_anno.Text = DateTime.Now.Date.Year.ToString();
                tx_anno.ReadOnly = true;

                tx_tipcam.Focus();
            }
            else this.Close();
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "EDICION";
            if (rb_pers.Checked == false && rb_omg.Checked == false)
            {
                rb_pers.Checked = true;
                rb_pers_Click(null, null);
            }
            limpiaObj();
            limpiaTE();
            escribe("EDICION");
            //Tx_fecha.ReadOnly = true;     // 31/01/2025
            //selecFecha1.Enabled = false;  // 31/01/2025
            //tx_tipcam.ReadOnly = true;    // 31/01/2025
            selecFecha1.Value = DateTime.Now.Date;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            pan_p.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            tx_idOper.Focus();
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.delete40;
            Tx_modo.Text = "BORRAR";
            if (rb_pers.Checked == false && rb_omg.Checked == false)
            {
                rb_pers.Checked = true;
                rb_pers_Click(null, null);
            }
            sololee("");
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            pan_p.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            tx_idOper.Focus();
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = null;
            Tx_modo.Text = "VISUALIZAR";
            if (rb_pers.Checked == false && rb_omg.Checked == false)
            {
                rb_pers.Checked = true;
                rb_pers_Click(null, null);
            }
            sololee("");
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            pan_p.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            tx_idOper.Focus();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "IMPRIMIR";
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

        #region limpiadores, readonlys
        private void limpiaObj()
        {
            OcatIn.codigo = "";                                       // Objeto categoría de ingreso
            OcatIn.nombre = "";
            OcatIn.largo = "";
            Omone.codigo = "";                                        // Objeto moneda
            Omone.nombre = "";
            Omone.siglas = "";
            Ocajd.codigo = "";                                        // Objeto cada de destino - desde donde sale el dinero
            Ocajd.nombre = "";
            Ocajd.largo = "";
            //Oprove.codigo = "";                                       // Objeto proveedor
            //Oprove.nombre = "";
            Omonto.codMOrige = "";                                    // Objeto monto
            Omonto.monDolar = 0;
            Omonto.monEuros = 0;
            Omonto.monOrige = 0;
            Omonto.monSoles = 0;
            Omonto.tipCDol = 0;
            Omonto.tipCOri = 0;
            //Ogiro.ctades = "";
            //Ogiro.tipodes = "";
            //Ogiro.idcod = "";
            //Ogiro.codigo = "";
            Oingreso.limpia();
        }
        private void limpiaTE() // limpia textbox, etiquetas, combos
        {
            tx_idOper.Clear();
            Tx_catIng.Clear();
            Tx_ctaDest.Clear();
            tx_ctaGiro.Clear();
            tx_descrip.Clear();
            tx_monto.Clear();
            //tx_tipcam.Clear();    05/09/2024 mejor no limpiamos el tipo de cambio
            tx_dat_giro.Clear();
            //
            eti_nomCaja.Text = "";
            eti_nomCat.Text = "";
            eti_nomCtaGiro.Text = "";
            chk_datSimil.Checked = false;
            cmb_mon.SelectedIndex = -1; // no puede ser 0 porque el objeto moneda esta limpio 02/09/2024
            chk_giroC.Checked = false;
        }
        private void escribe(string quien)  // pones los campos necesarios en readonly = false
        {
            if (quien == "EDICION") tx_idOper.ReadOnly = true;
            Tx_fecha.ReadOnly = false;
            Tx_catIng.ReadOnly = false;
            Tx_ctaDest.ReadOnly = false;
            tx_ctaGiro.ReadOnly = false;
            tx_descrip.ReadOnly = false;
            tx_monto.ReadOnly = false;
            tx_tipcam.ReadOnly = false;
            //
            cmb_mon.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            chk_datSimil.Enabled = true;
            chk_giroC.Enabled = true;
            cmb_mon.Enabled = true;
            cmb_mon.SelectedIndex = -1;
            cmb_mon_SelectedIndexChanged(null, null);
        }
        private void sololee(string quien)  //    // T=todos los campos, "" ó "C" campos comunes
        {
            Tx_catIng.ReadOnly = true;
            Tx_ctaDest.ReadOnly = true;
            tx_ctaGiro.ReadOnly = true;
            tx_descrip.ReadOnly = true;
            tx_monto.ReadOnly = true;
            tx_tipcam.ReadOnly = true;
            tx_idOper.ReadOnly = false;
            rb_omg.Enabled = false;
            rb_pers.Enabled = false;
            chk_datSimil.Enabled = false;
            chk_giroC.Enabled = false;
            cmb_mon.Enabled = false;
            if (quien == "T")
            {
                tx_idOper.ReadOnly = true;
                tx_anno.ReadOnly = true;
            }
        }
        #endregion

        #region radiobotones y checks
        private void rb_omg_Click(object sender, EventArgs e)
        {
            if (rb_omg.Checked == true)
            {
                eti_tituloForm.Text = eti_tituloForm.Tag.ToString() + "DE CUENTAS OMG";
                pan_p.Tag = "omg";
                limpiaTE();
                jalaGrilla(int.Parse(tx_diasA.Text), "cassaomg");  // jalaGrilla(diasAtroya, "cassaomg") 11/05/2025
                Tx_ctaDest.Values = lista_DES.ToArray();
                tx_ctaGiro.Values = lista_DES.ToArray();
            }
        }
        private void rb_pers_Click(object sender, EventArgs e)
        {
            if (rb_pers.Checked == true)
            {
                eti_tituloForm.Text = eti_tituloForm.Tag.ToString() + "DE CUENTAS PERSONALES";
                pan_p.Tag = "personal";
                limpiaTE();
                jalaGrilla(int.Parse(tx_diasA.Text), "cassaconti");  // jalaGrilla(diasAtroya, "cassaconti") 11/05/2025
                Tx_ctaDest.Values = lista_CON.ToArray();
                tx_ctaGiro.Values = lista_CON.ToArray();
            }
        }
        private void chk_giroC_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_giroC.CheckState == CheckState.Checked)
            {
                //tx_ctaGiro.Visible = true;
                //eti_nomCtaGiro.Visible = true;
            }
            else
            {
                //tx_ctaGiro.Visible = false;
                //eti_nomCtaGiro.Visible = false;
            }
        }
        private void chk_datSimil_CheckStateChanged(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO")
            {
                if (chk_datSimil.CheckState == CheckState.Checked)
                {
                    // Si los campos principales estan en blanco, jalamos el ultimo del dia,casa y tipo
                    if (Tx_catIng.Text == "" && Tx_ctaDest.Text == "")
                    {
                        //if(advancedDataGridView1.Rows.Count > 0) datsimil(); ... ya no 31/01/2025
                    }
                    else
                    {
                        // si los campos principales no estan en blanco, no jala nada 
                    }
                }

            }
        }
        #endregion

        #region enters, leaves y validaciones
        private void Tx_catIngre_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (Tx_catIng.Text.Trim() != "")
                    {
                        DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catIng.Text.Trim() + "'");
                        if (nc.Length > 0)  // "idtabella='CAM' and descrizionerid='" + Tx_catIng.Text.Trim() + "'");
                        {
                            eti_nomCat.Text = nc[0].ItemArray[3].ToString();    // ItemArray[2]
                            OcatIn.codigo = nc[0].ItemArray[1].ToString();      // ItemArray[1]
                            OcatIn.nombre = eti_nomCat.Text; // Tx_catIng.Text
                            OcatIn.largo = Tx_catIng.Text;  // eti_nomCat.Text
                        }
                        else
                        {
                            Tx_catIng.Clear();
                            eti_nomCat.Text = "";
                            MessageBox.Show("No existe el nombre del egreso");
                        }
                    }
                }
            }
        }
        private void Tx_catIngre_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_catIng.Text.Trim() != "")
                {
                    DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catIng.Text.Trim() + "' and numero=1");
                    if (nc.Length > 0)
                    {
                        eti_nomCat.Text = nc[0].ItemArray[3].ToString();    // nc[0].ItemArray[2].ToString()
                        OcatIn.codigo = nc[0].ItemArray[1].ToString();
                        OcatIn.largo = Tx_catIng.Text;    // OcatEg.nombre = Tx_catEgre.Text
                        OcatIn.nombre = eti_nomCat.Text;     // OcatEg.largo = eti_nomCat.Text
                    }
                    else
                    {
                        Tx_catIng.Clear();
                        eti_nomCat.Text = "";
                        MessageBox.Show("No existe el nombre del egreso");
                    }
                }
            }
        }
        private void Tx_ctaDes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (Tx_ctaDest.Text.Trim() != "" && Tx_ctaDest.Text.Length >= 3)  // *************** 14/12/2024)
                    {
                        xxx();
                        bool v = tx_descrip.Focus();
                        if (v == true) Conti3.AutoClosingMessageBox.Show("Ingrese la descripción", "", 10); // MessageBox.Show("Ingrese una descripción");
                    }
                }
            }
        }
        private void Tx_ctaDes_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_ctaDest.Text.Trim() != "" && Tx_ctaDest.Text.Length >= 3)
                {
                    string[] vuelto = oFEgres.ValiCtaCon(Tx_ctaDest.Text, (rb_omg.Checked == true) ? "OMG" : "PER", "algo");
                    if (vuelto.Length > 0 && vuelto[0] != "")
                    {
                        Ocajd.codigo = vuelto[0];
                        Ocajd.nombre = vuelto[1];
                        Ocajd.largo = vuelto[2];
                        eti_nomCaja.Text = Ocajd.nombre; //  = Ocajd.largo
                    }
                    else
                    {
                        Tx_ctaDest.Clear();
                        eti_nomCaja.Text = "";
                        MessageBox.Show("No existe el nombre de la cuenta");
                    }
                }
            }
        }
        private void tx_idOper_Validating(object sender, CancelEventArgs e)
        {
            if (tx_idOper.Text.Trim() != "" && !("NUEVO").Contains(Tx_modo.Text))
            {
                string[] retu = oFEgres.ValiIdOper((rb_omg.Checked == true) ? "OMG" : "PER", tx_idOper.Text.Trim(), tx_anno.Text, "E");
                if (retu[0] == "")
                {
                    limpiaObj();
                    limpiaTE();
                    MessageBox.Show("No existe el código de operación");
                }
                else
                {
                    // asignamos los valores de retu[] a los objetos
                    string anOp = "";
                    string fecOp = "";              // fecha de operacion
                    decimal tipca = 0;              // tip cambio del monto origen
                    string descr = "";              // descripcion de la operacion
                    string idmov = "";              // id del movimiento
                    if (rb_omg.Checked == true)
                    {
                        // ANNO,ID_MOVIM,FECHA,DESTINO,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                        //   0      1      2      3       4       5     6        7          8          9       10      11        12
                        // IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                        //     13         14        15     16       17        18           19            20          21      22
                        // CASA,GIRO_CTA,a.idgiroconto,CTA_DESTINO
                        //  23      24          25          26
                        anOp = retu[0];
                        fecOp = retu[2].Substring(0, 10);       // fecha
                        OcatIn.codigo = retu[14];               // IDCategoria
                        OcatIn.nombre = retu[4];                // INGRESO
                        OcatIn.largo = retu[19];                // DET_INGRESO
                        Omone.codigo = retu[15];                // "codimon"
                        Omone.siglas = retu[5];                 // "MONEDA"
                        Omone.nombre = retu[16];                // "nombmon"
                        Omonto.codMOrige = retu[15];            // "codimon"
                        Omonto.monOrige = decimal.Parse(retu[6]);   // "MONTO"
                        Omonto.tipCOri = decimal.Parse(retu[8]);   // "TCMonOri"
                        Omonto.monDolar = decimal.Parse(retu[11]);  // "ImportoDU"
                        Omonto.tipCDol = decimal.Parse(retu[8]);    // "TIP_CAMBIO"
                        Omonto.monSoles = decimal.Parse(retu[12]);  // "ImportoSU"
                        tipca = decimal.Parse(retu[8]);            // "TCMonOri"
                        Ocajd.codigo = retu[13];                // "IDDestino"
                        Ocajd.nombre = retu[3];                 // "DESTINO"
                        Ocajd.largo = retu[18];                 // "DET_DESTINO"
                        descr = retu[7];                        // "DESCRIPCION"
                        idmov = retu[1];                        // "ID_MOVIM"
                    }
                    else
                    {
                        // ANNO,ID_MOVIM,FECHA,CUENTA,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                        //  0       1      2     3       4       5     6        7          8           9      10       11        12
                        // IDConto,a.IDCategoria,a.codimon,a.nombmon,a.TCMonOri,DET_CUENTA,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                        //    13         14          15        16         17        18         19             20         21       22
                        // CASA,GIRO_CTA,a.IDGiroConto,CTA_DESTINO
                        //  23     24          25           26
                        anOp = retu[0];
                        fecOp = retu[2].Substring(0, 10);       // "FECHA"
                        OcatIn.codigo = retu[14];               // aca debe ser id del ingreso
                        OcatIn.nombre = retu[4];                // aca debe ser ingresos
                        OcatIn.largo = retu[19];                // aca tambien debe ser ingresos
                        Omone.codigo = retu[15];                // "codimon"
                        Omone.siglas = retu[5];                 // "MONEDA"
                        Omone.nombre = retu[16];                // "nombmon"
                        Omonto.codMOrige = retu[15];            // "codimon"
                        Omonto.monOrige = decimal.Parse(retu[6]);   // "MONTO"
                        Omonto.tipCOri = decimal.Parse(retu[8]);   // "TCMonOri"
                        Omonto.monDolar = decimal.Parse(retu[11]);  // "ImportoDU"
                        Omonto.tipCDol = decimal.Parse(retu[8]);    // "TIP_CAMBIO"
                        Omonto.monSoles = decimal.Parse(retu[12]);  // "ImportoSU"
                        tipca = decimal.Parse(retu[8]);        // "TCMonOri"
                        Ocajd.codigo = retu[13];                // "IDConto"
                        Ocajd.nombre = retu[3];                 // "CUENTA"
                        Ocajd.largo = retu[18];                 // "DET_CUENTA"
                        descr = retu[7];                        // "DESCRIPCION"
                        idmov = retu[1];                        // "ID_MOVIM"
                    }
                    Oingreso.creaIngreso(pan_p.Tag.ToString(), fecOp, OcatIn, Omone, Omonto, tipca,
                            Ocajd, descr, idmov, Ogiro, anOp);
                    jalaoc();
                }
            }
        }     // busca en toda la base de datos 
        #region no se usa
        /* private void tx_ctaGiro2_KeyPress(object sender, KeyPressEventArgs e)
{
    if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
    {
        if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
        {
            if (tx_ctaGiro.Text.Trim() != "")
            {
                DataRow[] row;
                if (rb_omg.Checked == true)
                {
                    row = Program.dt_definic.Select("idtabella='DES' and descrizione='" + tx_ctaGiro.Text.Trim() + "'");
                }
                else
                {
                    row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + tx_ctaGiro.Text.Trim() + "'");
                }
                if (row.Length > 0)
                {
                    foreach (DataRow dat in row)
                    {
                        tx_dat_giro.Text = dat[1].ToString();
                        eti_nomCtaGiro.Text = dat[3].ToString();
                        Ogiro.tipodes = (rb_omg.Checked == true) ? "OMG" : "PER";
                        Ogiro.ctades = eti_nomCtaGiro.Text; // tx_dat_giro.Text;
                        Ogiro.idcod = tx_dat_giro.Text;
                        Ogiro.largo = tx_ctaGiro.Text;  //eti_nomCtaGiro.Text
                    }
                }
                else
                {
                    tx_dat_giro.Clear();
                    eti_nomCtaGiro.Text = "";
                    Ogiro.tipodes = "";
                    Ogiro.ctades = "";
                    Ogiro.idcod = "";
                    Ogiro.largo = "";
                    Ogiro.codigo = "";
                }
            }
        }
    }
}   */
        #endregion
        private void tx_monto_Validating(object sender, CancelEventArgs e)
        {
            decimal monti = 0; decimal cambi = 0;
            decimal.TryParse(tx_monto.Text, out monti);
            tx_monto.Text = Math.Round(monti, 2).ToString("#,##0.00");
            decimal.TryParse(tx_tipcam.Text, out cambi);
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")   // Tx_modo.Text == "NUEVO" && monti > 0
            {
                Omonto.monOrige = monti;
                if (Omone.codigo == codDol)
                {
                    Omonto.tipCDol = cambi; // tcDia.tcD;
                    Omonto.tipCOri = cambi; // tcDia.tcD;
                    Omonto.monEuros = 0;
                    Omonto.monDolar = decimal.Parse(tx_monto.Text);
                    Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                    oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), decimal.Parse(tx_tipcam.Text));
                }
                if (Omone.codigo == codSol)
                {
                    Omonto.tipCDol = cambi; // tcDia.tcD;
                    Omonto.tipCOri = cambi; // tcDia.tcD;
                    Omonto.monEuros = 0;
                    Omonto.monSoles = decimal.Parse(tx_monto.Text); // Omonto.monDolar * Omonto.tipCOri;
                    Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                    oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCDol);
                }
                if (Omone.codigo == codEur)
                {
                    Omonto.tipCDol = 0;
                    Omonto.tipCOri = tcDia.tcE;
                    Omonto.monEuros = decimal.Parse(tx_monto.Text);
                    Omonto.monDolar = 0;
                    Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                    oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCOri);
                }
            }
        }
        private void tx_tipcam_Validating(object sender, CancelEventArgs e)
        {
            decimal monti = 0; decimal cambi = 0;
            decimal.TryParse(tx_monto.Text, out monti);
            decimal.TryParse(tx_tipcam.Text, out cambi);
            tx_tipcam.Text = Math.Round(cambi, 3).ToString("#0.000");
            if ((Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION") && monti > 0)
            {
                Omonto.monOrige = monti;
                if (true)
                {
                    if (Omonto.codMOrige == codDol) Omonto = oFEgres.calc_monedas(cmb_mon, monti, cambi);
                    if (Omonto.codMOrige == codSol) Omonto = oFEgres.calc_monedas(cmb_mon, monti, cambi);
                    if (Omonto.codMOrige == codEur) Omonto = oFEgres.calc_monedas(cmb_mon, monti, Omonto.tipCOri);
                }
            }
        }
        private void tx_tipcam_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO")   // 20/05/2025
            {
                //tx_idOper.Focus();
                Tx_fecha.Focus();
            }
        }
        private void selecFecha1_ValueChanged(object sender, EventArgs e)
        {
            // En ningun caso la fecha puede ser posterior al actual
            // si es nuevo la fecha puede ser anterior
            // si es edicion no se permite cambiar la fecha, 04/12/2024
            if (selecFecha1.Value.Date > DateTime.Now.Date)
            {
                MessageBox.Show("No se permite fechas posteriores","Atención",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                selecFecha1.Value = DateTime.Now.Date;
                Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
            }
            if ((Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION") && chk_datSimil.Checked == false)
            {
                Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
                tipCambio(null);
            }
        }
        private void Tx_fecha_Click(object sender, EventArgs e)
        {
            var mtb = (MaskedTextBox)sender;
            mtb.Select(0, 0);
            mtb.Focus();
        }
        private void Tx_fecha_Validating(object sender, CancelEventArgs e)
        {
            // En ningun caso la fecha puede ser posterior al actual
            // si es nuevo la fecha puede ser anterior
            // si es edicion no se permite cambiar la fecha, 04/12/2024
            try
            {
                if (Tx_fecha.Text.Length != 10) Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
                DateTime fecOp = DateTime.Parse(Tx_fecha.Text);
                if (fecOp > DateTime.Now.Date)
                {
                    MessageBox.Show("No se permite fechas posteriores", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    selecFecha1.Value = DateTime.Now.Date;
                    Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
                }
            }
            catch (Exception ex)
            {
                DateTime fecOp = DateTime.Now.Date;
                Tx_fecha.Text = fecOp.ToString();
            }
        }
        private void TextBox_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
            /*
            if (textBox.Name == "Tx_catIng")
            {
                DisableMouse();
            } */
        }
        private void xxx()
        {
            if (Tx_ctaDest.Text.Trim() != "")
            {
                DataRow[] row;
                if (rb_omg.Checked == true)
                {
                    row = Program.dt_definic.Select("idtabella='DES' and descrizione='" + Tx_ctaDest.Text.Trim() + "'");
                }
                else
                {
                    row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + Tx_ctaDest.Text.Trim() + "'");
                }
                if (row.Length > 0)
                {
                    eti_nomCaja.Text = row[0].ItemArray[3].ToString().ToUpper();  //
                    Ocajd.codigo = row[0].ItemArray[1].ToString().ToUpper();
                    Ocajd.nombre = eti_nomCaja.Text;    // Tx_ctaDest.Text;
                    Ocajd.largo = Tx_ctaDest.Text.ToUpper();      // eti_nomCaja.Text

                    tx_descrip.Focus();
                }
                else
                {
                    Tx_ctaDest.Clear();
                    eti_nomCaja.Text = "";
                    MessageBox.Show("No existe el nombre de la cuenta");
                }
            }
        }
        private void tx_descrip_Enter(object sender, EventArgs e)
        {
            // sino colocamos esto se va a autoseleccionar todo el texto y al dar <enter> se borrará
            //tx_descrip.SelectionStart = tx_descrip.Text.Length;   // ya no ... 24/01/2025
        }       // funcion para evitar que se autoseleccione todo el texto del campo
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        #endregion

        #region combos
        private void cmb_mon_SelectedValueChanged(object sender, EventArgs e)
        {
            // se copio igualito a changecommitted
        }   // selección de moneda
        private void cmb_mon_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
            if (cmb_mon.SelectedIndex > -1 && (cmb_mon.SelectedValue != null && cmb_mon.SelectedValue.ToString() != ""))
            {
                Omone.codigo = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                Omone.siglas = cmb_mon.Text;    // siglas de la moneda
                Omonto.codMOrige = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                DataRow[] row = Program.dt_definic.Select("idtabella='MON' and idcodice='" + Omone.codigo + "'");
                if (row.Length > 1) Omone.nombre = row[0].ItemArray[2].ToString();
                if (tx_monto.Text != "" && tx_tipcam.Text != "")
                {
                    Omonto.monOrige = decimal.Parse(tx_monto.Text);
                    if (Omone.codigo == codDol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monDolar = decimal.Parse(tx_monto.Text);
                        Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                        oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), decimal.Parse(tx_tipcam.Text));
                    }
                    if (Omone.codigo == codSol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monSoles = decimal.Parse(tx_monto.Text); // Omonto.monDolar * Omonto.tipCOri;
                        Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                        oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCDol);
                    }
                    if (Omone.codigo == codEur)
                    {
                        Omonto.tipCDol = 0;
                        Omonto.tipCOri = tcDia.tcE;
                        Omonto.monEuros = decimal.Parse(tx_monto.Text);
                        Omonto.monDolar = 0;
                        Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                        oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCOri);
                    }
                }
            }   */
        }
        private void cmb_mon_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "" && (cmb_mon.SelectedValue != null && cmb_mon.SelectedValue.ToString() != ""))
            {
                Omone.codigo = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                Omone.siglas = cmb_mon.Text;    // siglas de la moneda
                Omonto.codMOrige = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                DataRow[] row = Program.dt_definic.Select("idtabella='MON' and idcodice='" + Omone.codigo + "'");
                Omone.nombre = row[0].ItemArray[2].ToString();
                if (tx_monto.Text != "" && tx_tipcam.Text != "")
                {
                    Omonto.monOrige = decimal.Parse(tx_monto.Text);
                    if (Omone.codigo == codDol)
                    {
                        Omonto.tipCDol = tcDia.tcD; // Omonto.tipCOri;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monDolar = decimal.Parse(tx_monto.Text);
                        Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                        oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), decimal.Parse(tx_tipcam.Text));
                    }
                    if (Omone.codigo == codSol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monSoles = decimal.Parse(tx_monto.Text); // Omonto.monDolar * Omonto.tipCOri;
                        Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                        oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCDol);
                    }
                    if (Omone.codigo == codEur)
                    {
                        Omonto.tipCDol = 0;
                        Omonto.tipCOri = tcDia.tcE;
                        Omonto.monEuros = decimal.Parse(tx_monto.Text);
                        Omonto.monDolar = 0;
                        Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                        oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCOri);
                    }
                }
            }
        }
        #endregion

        #region datagridview
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (true)    // 24/01/2025 (Tx_modo.Text != "NUEVO")
            {
                string annOp = "";
                string fecOp = "";              // fecha de operacion
                decimal tipca = 0;              // tip cambio del monto origen
                string descr = "";              // descripcion de la operacion
                string idmov = "";              // id del movimiento
                if (rb_omg.Checked == true)
                {
                    // CASA,AÑO,ID_MOVIM,FECHA,DESTINO,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,GIRO_CTA,idgiroconto,CTA_DESTINO,
                    // usuario,dia,ImportoDE,ImportoSE,IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_INGRESO,CodGiro
                    if (Tx_modo.Text != "NUEVO") fecOp = advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);   // 24/01/2025
                    OcatIn.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDCategoria"].Value.ToString();
                    OcatIn.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["INGRESO"].Value.ToString();
                    OcatIn.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_INGRESO"].Value.ToString();    // debe ser ingreso
                    Omone.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = advancedDataGridView1.Rows[e.RowIndex].Cells["MONEDA"].Value.ToString();
                    Omone.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["MONTO"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoDE"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoSE"].Value.ToString());
                    tipca = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Ocajd.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDDestino"].Value.ToString();
                    Ocajd.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["DESTINO"].Value.ToString();
                    Ocajd.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_DESTINO"].Value.ToString();
                    descr = advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    if (Tx_modo.Text != "NUEVO") idmov = advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString();     // 24/01/2025
                    annOp = advancedDataGridView1.Rows[e.RowIndex].Cells["ANNO"].Value.ToString();
                }
                else
                {
                    // CASA,AÑO,ID_MOVIM,FECHA,CUENTA,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,GIRO_CTA,IDGiroConto,CTA_DESTINO,
                    // usuario,dia,ImportoDE,ImportoSE,IDConto,IDCategoria,codimon,nombmon,TCMonOri,DET_CUENTA,DET_INGRESO,CodGiro 
                    if (Tx_modo.Text != "NUEVO") fecOp = advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);   // 24/01/2025
                    OcatIn.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDCategoria"].Value.ToString();
                    OcatIn.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["INGRESO"].Value.ToString();
                    OcatIn.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_INGRESO"].Value.ToString(); 
                    Omone.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = advancedDataGridView1.Rows[e.RowIndex].Cells["MONEDA"].Value.ToString();
                    Omone.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["MONTO"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoDE"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoSE"].Value.ToString());
                    tipca = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Ocajd.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDConto"].Value.ToString();
                    Ocajd.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["CUENTA"].Value.ToString();
                    Ocajd.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_CUENTA"].Value.ToString();
                    descr = advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    if (Tx_modo.Text != "NUEVO") idmov = advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString();     // 24/01/2025
                    annOp = advancedDataGridView1.Rows[e.RowIndex].Cells["ANNO"].Value.ToString();
                }
                Oingreso.creaIngreso(pan_p.Tag.ToString(), fecOp, OcatIn, Omone, Omonto, tipca,
                        Ocajd, descr, idmov, Ogiro, annOp);
                jalaoc();
            }
        }
        private void insFilaEnDataG(string _casa, string _corre)
        {
            DataRow fila = dt_grillaI.NewRow();
            string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
            advancedDataGridView1.Rows[0].DefaultCellStyle.BackColor = System.Drawing.SystemColors.Window;  // 22/04/2025
            if (rb_omg.Checked == true)
            {
                // CASA,AÑO,ID_MOVIM,FECHA,DESTINO,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,idgiroconto,CTA_DESTINO,
                // usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDDestino,IDCategoria
                // , , Omone, Omonto, decimal.Parse(tx_tipcam.Text), Ocajd, Oprove, tx_descrip.Text, corre
                fila["CASA"] = _casa;
                fila["ANNO"] = tx_anno.Text;
                fila["ID_MOVIM"] = _corre;
                fila["FECHA"] = fecOp;
                fila["DESTINO"] = Ocajd.nombre;     // nombre cuenta destino
                fila["INGRESO"] = OcatIn.nombre.ToUpper();     // nombre categoria egreso
                fila["MONEDA"] = Omone.siglas;      // siglas moneda origen
                fila["MONTO"] = Omonto.monOrige;    // valor origen
                fila["DESCRIPCION"] = tx_descrip.Text;
                fila["TIP_CAMBIO"] = Omonto.tipCOri; // decimal.Parse(tx_tipcam.Text);
                //fila["PROVEEDOR"] = Oprove.nombre;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["idgiroconto"] = Ogiro.idcod;
                fila["CTA_DESTINO"] = Ogiro.largo;
                fila["usuario"] = Program.vg_user;
                //fila["dia"] = "";
                fila["ImportoDE"] = Omonto.monDolar;
                fila["ImportoSE"] = Omonto.monSoles;
                //fila["idanagrafica"] = Oprove.codigo;
                fila["IDDestino"] = Ocajd.codigo;
                fila["IDCategoria"] = OcatIn.codigo;
                fila["codimon"] = Omone.codigo;
                fila["nombmon"] = Omone.nombre;
                fila["TCMonOri"] = tx_tipcam.Text;
                fila["DET_DESTINO"] = Ocajd.largo;
                fila["DET_INGRESO"] = OcatIn.largo;
                fila["tipodesgiro"] = Ogiro.tipodes;
                fila["CodGiro"] = Ogiro.codigo;
                fila["CTA_GIRO"] = Ogiro.ctades;
            }
            if (rb_pers.Checked == true)
            {
                // CASA,AÑO,ID_MOVIM,FECHA,CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,a.IDGiroConto,CTA_DESTINO,
                // usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,codimon,nombmon,TCMonOri
                fila["CASA"] = _casa;
                fila["ANNO"] = tx_anno.Text;
                fila["ID_MOVIM"] = _corre;
                fila["FECHA"] = fecOp;
                fila["CUENTA"] = Ocajd.nombre;
                fila["INGRESO"] = OcatIn.nombre.ToUpper();
                fila["MONEDA"] = Omone.siglas;
                fila["MONTO"] = Omonto.monOrige;
                fila["DESCRIPCION"] = tx_descrip.Text;
                fila["TIP_CAMBIO"] = Omonto.tipCOri; // decimal.Parse(tx_tipcam.Text);
                //fila["PROVEEDOR"] = Oprove.nombre;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["IDGiroConto"] = Ogiro.idcod;
                fila["CTA_DESTINO"] = Ogiro.largo;
                fila["usuario"] = Program.vg_user;
                //fila["dia"] = "";
                fila["ImportoDE"] = Omonto.monDolar;
                fila["ImportoSE"] = Omonto.monSoles;
                //fila["idanagrafica"] = Oprove.codigo;
                fila["IDConto"] = Ocajd.codigo;
                fila["IDCategoria"] = OcatIn.codigo;
                fila["CTA_GIRO"] = Ogiro.ctades;
                fila["codimon"] = Omone.codigo;
                fila["nombmon"] = Omone.nombre;
                fila["TCMonOri"] = Omonto.tipCOri;
                fila["DET_CUENTA"] = Ocajd.largo;
                fila["DET_INGRESO"] = OcatIn.largo;
                fila["tipodesgiro"] = Ogiro.tipodes;
                fila["CodGiro"] = Ogiro.codigo;
            }
            dt_grillaI.Rows.InsertAt(fila, 0);
            advancedDataGridView1.CurrentCell = advancedDataGridView1.Rows[0].Cells[0];
            advancedDataGridView1.CurrentRow.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(col1rafila);   // 22/04/2025
        }                                           // INSERTA en la grilla el registro nuevo despues de grabar en la B.D.
        private void jalaGrilla(int dAtras, string ntabla)
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        string consulta = "";
                        if (ntabla == "cassaomg")
                        {
                            consulta = "ConIngre_cassaOmg";
                        }
                        if (ntabla == "cassaconti")
                        {
                            consulta = "ConIngre_cassaConti";
                        }
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@Vdias", dAtras);
                            micon.Parameters.AddWithValue("@Vanno", 0);
                            micon.Parameters.AddWithValue("@Vidmov", "");
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_grillaI.Clear();
                                dt_grillaI.Columns.Clear();
                                da.Fill(dt_grillaI);
                                advancedDataGridView1.DataSource = dt_grillaI;
                            }
                        }
                        armaGrilla(advancedDataGridView1, limCols);      // cuadramos las columnas de la grilla
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error de conexión al servidor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Application.Exit();
                }
            }
        }                      // muestra datos de la fecha actual hasta <dAtras> días atras 
        private void armaGrilla(AdvancedDataGridView dgv_, int filasLim) // DataGridView dgv_, int filasLim
        {
            if (dgv_.Rows.Count > 1)
            {
                for (int i = 0; i < dgv_.Columns.Count; i++)
                {
                    if (i > filasLim)
                    {
                        dgv_.Columns[i].Visible = false;
                    }
                    else
                    {
                        dgv_.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        _ = decimal.TryParse(dgv_.Rows[0].Cells[i].Value.ToString(), out decimal vd);
                        if (vd != 0) dgv_.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                }
                int b = 0;
                for (int i = 0; i < dgv_.Columns.Count; i++)
                {
                    int a = dgv_.Columns[i].Width;
                    b += a;
                    dgv_.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    dgv_.Columns[i].Width = a;
                }
                if (b < dgv_.Width) dgv_.Width = b - 20;
                dgv_.ReadOnly = true;
            }
            else
            {

            }
        }                 // ajusta el ancho de las columnas y muestra hasta el limite
        public void actFilaEnDataI(DataTable dt, string _casa, string _corre)
        {
            string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = dt.Rows[i];
                if (dr["ANNO"].ToString() == tx_anno.Text && dr["ID_MOVIM"].ToString() == oFEgres.CDerecha(_corre, 6))
                {
                    if (rb_omg.Checked == true)
                    {
                        // ANNO,ID_MOVIM,FECHA,DESTINO,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                        // IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                        // CASA,GIRO_CTA,a.idgiroconto,CTA_DESTINO
                        dr["CASA"] = _casa;
                        dr["ANNO"] = tx_anno.Text;
                        dr["ID_MOVIM"] = _corre;
                        dr["FECHA"] = fecOp;
                        dr["DESTINO"] = Ocajd.nombre;     // nombre cuenta destino
                        dr["INGRESO"] = OcatIn.nombre.ToUpper();     // nombre categoria ingreso
                        dr["MONEDA"] = Omone.siglas;      // siglas moneda origen
                        dr["MONTO"] = Omonto.monOrige;    // valor origen
                        dr["DESCRIPCION"] = tx_descrip.Text;
                        dr["TIP_CAMBIO"] = Omonto.tipCOri;// decimal.Parse(tx_tipcam.Text);
                        //dr["PROVEEDOR"] = ;
                        dr["GIRO_CTA"] = "";    // Ogiro.tipodes;
                        dr["idgiroconto"] = ""; // Ogiro.idcod;
                        dr["CTA_DESTINO"] = "";  //;
                        dr["usuario"] = Program.vg_user;
                        //dr["dia"] = "";
                        dr["ImportoDE"] = Omonto.monDolar;
                        dr["ImportoSE"] = Omonto.monSoles;
                        //dr["idanagrafica"] = ;
                        dr["IDDestino"] = Ocajd.codigo;
                        dr["IDCategoria"] = OcatIn.codigo;
                        dr["codimon"] = Omone.codigo;
                        dr["nombmon"] = Omone.nombre;
                        dr["TCMonOri"] = Omonto.tipCOri;    // tx_tipcam.Text;
                        dr["DET_DESTINO"] = Ocajd.largo;
                        dr["DET_INGRESO"] = OcatIn.largo;
                        dr["tipodesgiro"] = ""; // Ogiro.tipodes;
                        dr["CodGiro"] = "";     // Ogiro.codigo;
                        dr["CTA_GIRO"] = "";    // Ogiro.ctades;
                    }
                    if (rb_pers.Checked == true)
                    {
                        // ANNO,ID_MOVIM,FECHA,CUENTA,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                        // IDConto,a.IDCategoria,a.codimon,a.nombmon,a.TCMonOri,DET_CUENTA,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                        // CASA,GIRO_CTA,a.IDGiroConto,CTA_DESTINO
                        dr["CASA"] = _casa;
                        dr["ANNO"] = tx_anno.Text;
                        dr["ID_MOVIM"] = _corre;
                        dr["FECHA"] = fecOp;
                        dr["CUENTA"] = Ocajd.nombre;
                        dr["INGRESO"] = OcatIn.nombre.ToUpper();
                        dr["MONEDA"] = Omone.siglas;
                        dr["MONTO"] = Omonto.monOrige;
                        dr["DESCRIPCION"] = tx_descrip.Text;
                        dr["TIP_CAMBIO"] = Omonto.tipCOri;// decimal.Parse(tx_tipcam.Text);
                        //dr["PROVEEDOR"] = ;
                        //dr["GIRO_CTA"] = Ogiro.tipodes;
                        //dr["IDGiroConto"] = Ogiro.idcod;
                        //dr["CTA_DESTINO"] = Ogiro.largo;
                        dr["usuario"] = Program.vg_user;
                        //dr["dia"] = "";
                        dr["ImportoDE"] = Omonto.monDolar;
                        dr["ImportoSE"] = Omonto.monSoles;
                        //dr["idanagrafica"] = ;
                        dr["IDConto"] = Ocajd.codigo;
                        dr["IDCategoria"] = OcatIn.codigo;
                        dr["codimon"] = Omone.codigo;
                        dr["nombmon"] = Omone.nombre;
                        dr["TCMonOri"] = Omonto.tipCOri;
                        dr["DET_CUENTA"] = Ocajd.largo;
                        dr["DET_INGRESO"] = OcatIn.largo;
                        dr["tipodesgiro"] = Ogiro.tipodes;
                        //dr["CodGiro"] = Ogiro.codigo;
                    }
                    dr.AcceptChanges();
                }
            }
            dt.AcceptChanges();
        }                // ACTUALIZA la grilla despues de haber actualizado la tabla
        private void advancedDataGridView1_SortStringChanged(object sender, EventArgs e)
        {
            DataTable dtg = (DataTable)advancedDataGridView1.DataSource;
            dtg.DefaultView.Sort = advancedDataGridView1.SortString;
        }
        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)                  // filtro de las columnas
        {
            DataTable dtg = (DataTable)advancedDataGridView1.DataSource;
            dtg.DefaultView.RowFilter = advancedDataGridView1.FilterString;
        }
        #endregion

        #region botones Grabar, nuevo prov.
        private void Bt_graba_Enter(object sender, EventArgs e)
        {
            Bt_graba.BackColor = Color.DarkSeaGreen; // ColorTranslator.FromHtml("#fabdba");
            if (Tx_modo.Text == "BORRAR") Bt_graba.Image = Conti3.Properties.Resources.delete40rojo;
            else Bt_graba.Image = Properties.Resources.save_item40_rojo;
        }
        private void Bt_graba_Leave(object sender, EventArgs e)
        {
            Bt_graba.BackColor = ColorTranslator.FromHtml("#667d97");
            if (Tx_modo.Text == "NUEVO") Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            if (Tx_modo.Text == "EDICION") Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            if (Tx_modo.Text == "BORRAR") Bt_graba.Image = Conti3.Properties.Resources.delete40;
        }
        private void Bt_graba_Click(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO")
            {
                // validamos datos esenciales
                if (Tx_catIng.Text == "")
                {
                    errorProvider1.SetError(Tx_catIng, "Debe ingresar un tipo");
                    Tx_catIng.Focus();
                    return;
                }
                if (Tx_ctaDest.Text == "")
                {
                    errorProvider1.SetError(Tx_ctaDest, "Debe seleccionar la cuenta");
                    Tx_ctaDest.Focus();
                    return;
                }
                errorProvider1.SetError(Tx_catIng, "");
                errorProvider1.SetError(Tx_ctaDest, "");
                if (cmb_mon.Text == "")
                {
                    errorProvider1.SetError(cmb_mon, "Debe seleccionar la moneda");
                    cmb_mon.Focus();
                    return;
                }
                errorProvider1.SetError(cmb_mon, "");
                if (tx_tipcam.Text.Trim() == "0" || tx_tipcam.Text.Trim() == "")
                {
                    errorProvider1.SetIconAlignment(tx_tipcam, ErrorIconAlignment.TopLeft);
                    errorProvider1.SetError(tx_tipcam, "Debe ingresar el tipo de cambio");
                    tx_tipcam.Focus();
                    return;
                }
                errorProvider1.SetError(tx_tipcam, "");
                if (tx_monto.Text == "")
                {
                    errorProvider1.SetError(tx_monto, "Debe ingresar un valor");
                    tx_monto.Focus();
                    return;
                }
                errorProvider1.SetError(tx_monto, "");
                // ******************* hoy 13/12/2024
                if (String.IsNullOrEmpty(OcatIn.codigo) || String.IsNullOrEmpty(OcatIn.nombre) || String.IsNullOrEmpty(OcatIn.largo))
                {
                    MessageBox.Show("Complete la categoría de ingreso", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_catIng.Focus();
                    return;
                }
                if (Tx_catIng.Text.Trim() != OcatIn.largo)
                {
                    MessageBox.Show("Categoría de ingreso no completa", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_catIng.Focus();
                    return;
                }
                if (String.IsNullOrEmpty(Ocajd.codigo) || String.IsNullOrEmpty(Ocajd.nombre) || String.IsNullOrEmpty(Ocajd.largo))
                {
                    MessageBox.Show("Complete la cuenta de destino", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_ctaDest.Text = "";
                    eti_nomCaja.Text = "";
                    Tx_ctaDest.Focus();
                    return;
                }
                if (Tx_ctaDest.Text.Trim() != Ocajd.largo)  // me quede aca ... Ocajd.largo debe estaría en UPPER
                {
                    MessageBox.Show("Cuenta destino no completa", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    //tx_descrip.Focus();
                    Tx_ctaDest.Select();
                    return;
                }
                // **********************************
                graba_nuevo();
                Tx_catIng.Focus();
            }
            if (Tx_modo.Text == "EDICION")
            {
                if (tx_idOper.Text == "")
                {
                    MessageBox.Show("No hay registro que Editar!", "Identificador en blanco", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }
                var aaa = MessageBox.Show("Confirma que desea EDITAR el Ingreso?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    graba_edicion();
                    limpiaObj();
                    limpiaTE();
                }
            }
            if (Tx_modo.Text == "BORRAR")
            {
                // validamos que exista registro que borrar
                if (tx_idOper.Text == "")
                {
                    MessageBox.Show("No hay registro que borrar!", "Identificador en blanco", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }
                var aaa = MessageBox.Show("Confirma que desea BORRAR el Ingreso?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    string tabla = "";
                    if (rb_omg.Checked == true) tabla = "cassaomg";
                    else tabla = "cassaconti";
                    //oFEgres.graba_borrar(tabla, selecFecha1.Value.Year.ToString(), "000000000" + oFEgres.CDerecha(tx_idOper.Text, 6), dt_grillaI);
                    oFEgres.graba_borrar(tabla, tx_anno.Text, "000000000" + oFEgres.CDerecha(tx_idOper.Text, 6), dt_grillaI);
                    limpiaObj();
                    limpiaTE();
                }
            }
        }
        private void bt_refresh_Click(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "")
            {
                jalaGrilla(int.Parse(tx_diasA.Text), (rb_omg.Checked == true) ? "cassaomg" : "cassaconti");
            }
        }
        private void graba_nuevo()
        {
            var aaa = MessageBox.Show("Confirma que desea crear el Ingreso?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (aaa == DialogResult.Yes)
            {
                string fecOp = Tx_fecha.Text;
                Ingresos Oingresos = new Ingresos();
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
                        if (Tx_modo.Text == "NUEVO")
                        {
                            string corre = oFEgres.correlativo(conn, ((rb_omg.Checked == true) ? "MCA" : "MCO"), int.Parse(tx_anno.Text)); // selecFecha1.Value.Date.Year
                            string corrA = corre;
                            if (corre != "error" && corre != "")
                            {
                                try
                                {
                                    Oingresos.creaIngreso(pan_p.Tag.ToString(), fecOp, OcatIn, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                                        Ocajd, tx_descrip.Text, corre, Ogiro, tx_anno.Text);
                                    Oingresos.grabaIngreso(conn);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Error en grabar el Ingreso");
                                    return;
                                }
                                insFilaEnDataG("LIM", oFEgres.CDerecha("00000" + corrA, 6));       // inserta el registro nuevo en la grilla
                                if (chk_datSimil.CheckState == CheckState.Checked)
                                {
                                    datsimil(); // 31/01/2025 ... solo deja los objetos y campos como estan
                                }
                                else
                                {
                                    limpiaObj();
                                    limpiaTE();
                                }
                            }
                            else
                            {
                                MessageBox.Show("Error en grabar los datos del ingreso", "No se completo la operación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }
        private void graba_edicion()
        {
            if (true)
            {
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
                        string fecOp = Tx_fecha.Text; // selecFecha1.Value.Date.ToShortDateString();
                        string corre = tx_anno.Text + tx_idOper.Text;
                        Oingreso.creaIngreso(pan_p.Tag.ToString(), fecOp, OcatIn, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                                        Ocajd, tx_descrip.Text, corre, Ogiro, tx_anno.Text);
                        Oingreso.EditaIngreso(conn, tx_anno.Text, "000000000" + oFEgres.CDerecha(tx_idOper.Text, 6));
                        actFilaEnDataI(dt_grillaI, "LIM", tx_idOper.Text);
                    }
                }
            }
        }

        #endregion

        private void tx_descrip_MouseClick(object sender, MouseEventArgs e)
        {
            tx_descrip.Focus();
        }

        private void Finan_Ingres_Click(object sender, EventArgs e)
        {
            this.Activate();
            this.BringToFront();
        }
    }
}
