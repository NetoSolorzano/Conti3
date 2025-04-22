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
    public partial class Finan_Egres : Form1
    {
        string nomForm = "Finan_Egres";
        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";port=" + login.port + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data +
            ";ConnectionLifeTime=" + login.ctl + ";";
        // datos de la grilla
        internal DataTable dt_grillaE = new DataTable();
        //
        DataTable dtpro = new DataTable();
        publicoConf conf = new publicoConf();
        List<string> lista_CAM = new List<string>();                                // categorias
        List<string> lista_DES = new List<string>();                                // cuentas DES
        List<string> lista_CON = new List<string>();                                // cuentas CON
        List<string> lista_prov = new List<string>();                               // lista de proveedores activos
        //
        tipcamDia tcDia = new tipcamDia();
        catEgresos OcatEg = new catEgresos();                                       // Objeto categoría de egreso
        monedas Omone = new monedas();                                              // Objeto moneda
        cajDestino Ocajd = new cajDestino();                                        // Objeto cada de destino - desde donde sale el dinero
        provees Oprove = new provees();                                             // Objeto proveedor
        montos Omonto = new montos();                                               // Objeto monto
        giroConto Ogiro = new giroConto();                                          // Objeto giroconto
        ccolores OColores = new ccolores();
        //
        Ingresos Oingresos = new Ingresos();
        Egresos Oegreso = new Egresos();
        int diasAtroya = 0;                                                         // dias atras hasta donde mostrará la grilla
        int limCols = 1;                                                            // limite de columnas que muestra la grilla
        string codDol = "MON002";
        string codEur = "MON003";
        string codSol = "MON001";
        string col1rafila = "";                                                     // color html de la 1ra fila en ingresos

        public Finan_Egres()
        {
            InitializeComponent();                  // inicializa los objetos graficos
            CargaINI(this);                         // colorea los objetos graficos
            CargaFormatos();                        // jala datos de combos y demas
            chk_giroC_CheckedChanged(null, null);   // 
            sololee("T");                           // T=todos los campos, "" ó "C" campos comunes
            jalainfo();                             // jala variables de tabla enlace
            initCampos();                           // pone maximos y upper case de campos texto
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                jalacolores(conn, OColores, nomForm);
                toolboton(conn);
            }
            colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);    // pinta el mundo de colores!
            tx_descrip.BackColor = ColorTranslator.FromHtml(OColores.Fondo_suave);
            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml(OColores.Fondo_boton_graba);
            Bt_graba.Image = null;
        }
        private void Finan_Egres_KeyDown(object sender, KeyEventArgs e)
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
                if (Tx_ctaDes.Focused == true)   // Tx_ctaDes.Focus() == true
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
                            Tx_ctaDes.Text = ayu2.ReturnValueA[2];      // ayu2.ReturnValueA[1]
                            eti_nomCaja.Text = ayu2.ReturnValueA[1];    // ayu2.ReturnValueA[2]
                            xxx();
                            //SendKeys.SendWait("{TAB}");
                            tx_provee.Focus();
                        }
                    }
                }
                if (tx_ctaGiro.Focused == true)
                {
                    para1 = "personal";  // (rb_omg.Checked == true) ? "omg" : "personal"
                    para2 = "cuenta";
                    para3 = "activos";    // todos | activos
                    ayuda2 ayu2 = new ayuda2(para1, para2, para3, para4);
                    var result = ayu2.ShowDialog();
                    if (result == DialogResult.Cancel)
                    {
                        if (!string.IsNullOrEmpty(ayu2.ReturnValue1))   // 0=codigo, 1=descripCorta, 2=descripLarga
                        {
                            tx_ctaGiro.Text = ayu2.ReturnValueA[2];
                            eti_nomCtaGiro.Text = ayu2.ReturnValueA[1];
                            tx_dat_giro.Text = ayu2.ReturnValueA[0];
                            yyy();
                        }
                    }
                }
                if (Tx_catEgre.Focused == true)
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
                            Tx_catEgre.Text = ayu2.ReturnValueA[2];// ayu2.ReturnValueA[1]
                            eti_nomCat.Text = ayu2.ReturnValueA[1];// ayu2.ReturnValueA[2]
                            OcatEg.codigo = ayu2.ReturnValueA[0];
                            OcatEg.nombre = ayu2.ReturnValueA[2];   // ayu2.ReturnValueA[1]
                            OcatEg.largo = ayu2.ReturnValueA[1];    // ayu2.ReturnValueA[2]
                        }
                    }
                }
                if (tx_prov.Focused == true)  // tx_provee.Focused == true
                {
                    para1 = "provee";
                    para2 = "";
                    para3 = "activos";    // todos | activos
                    ayuda2 ayu2 = new ayuda2(para1, para2, para3, para4);
                    var result = ayu2.ShowDialog();
                    if (result == DialogResult.Cancel)
                    {
                        if (!string.IsNullOrEmpty(ayu2.ReturnValue1))   // 0=codigo, 1=descripCorta, 2=descripLarga
                        {
                            tx_dat_provee.Text = ayu2.ReturnValueA[0];
                            Tx_nomProv.Text = ayu2.ReturnValueA[1];
                            tx_prov.Text = ayu2.ReturnValueA[1];
                            eti_nomprovee.Text = ayu2.ReturnValueA[1]; // 31/08/2024 ya no usamos
                            Oprove.codigo = tx_dat_provee.Text;
                            Oprove.nombre = Tx_nomProv.Text;
                            SendKeys.Send("{Tab}");
                        }
                    }
                }
                return true;    // indicate that you handled this keystroke
            }
            // Call the base class
            return base.ProcessCmdKey(ref msg, keyData);
        }    // F1 
        private void CargaFormatos()
        {
            // proveedores
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message,"Error de conexión",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    Application.Exit();
                    return;
                }
                string consulta = "SELECT idanagrafica,trim(upper(ragionesociale)) AS nombre FROM anag_for WHERE stato=1";
                using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                {
                    using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                    {
                        da.Fill(dtpro);
                    }
                }
                foreach (DataRow row in dtpro.Rows)
                {
                    lista_prov.Add(row["nombre"].ToString());
                }
                Tx_nomProv.Values = lista_prov.ToArray();
            }
            // categorias
            DataRow[] depar = Program.dt_definic.Select("idtabella='CAM' and numero=1");
            foreach (DataRow row in depar)
            {
                lista_CAM.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_catEgre.Values = lista_CAM.ToArray();
            // cuentas OMG
            depar = Program.dt_definic.Select("idtabella='DES' and numero=1");
            foreach (DataRow row in depar)
            {
                lista_DES.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            // cuentas personales
            depar = Program.dt_definic.Select("idtabella='CON' and numero=1");
            foreach (DataRow row in depar)
            {
                lista_CON.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            // monedas
            depar = Program.dt_definic.Select("idtabella='MON' and numero=1");
            cmb_mon.DataSource = depar.CopyToDataTable();
            cmb_mon.DisplayMember = "descrizionerid";
            cmb_mon.ValueMember = "idcodice";
        }
        public void jalacolores(MySqlConnection conn, ccolores OColores, string formu)
        {
            string consu = "select property,uvalue from dtproperties where objectid=@oid and value=@forma";
            using (MySqlCommand micon = new MySqlCommand(consu, conn))
            {
                micon.Parameters.AddWithValue("@oid", "COLOR");
                micon.Parameters.AddWithValue("@forma", formu);
                using (MySqlDataReader dr = micon.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        if (dr["property"].ToString() == "Fondo_fuerte") OColores.Fondo_fuerte = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Fondo_suave") OColores.Fondo_suave = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Fondo_normal") OColores.Fondo_normal = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Fondo_boton_graba") OColores.Fondo_boton_graba = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Fondo_pageFrame") OColores.Fondo_pageFrame = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Fondo_grilla") OColores.Fondo_grilla = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Grilla_fila_normal") OColores.Grilla_fila_normal = dr["uvalue"].ToString();
                        if (dr["property"].ToString() == "Grilla_fila_anulada") OColores.Grilla_fila_anulada = dr["uvalue"].ToString();
                    }
                }
            }
        }
        public void colorea(Form este, string fuerte, string normal, string suave)
        {
            este.BackColor = ColorTranslator.FromHtml(normal); // cuando usamos FromHtml NO da error por fondo transparente
            foreach (System.Windows.Forms.Control oControl in este.Controls)
            {
                if (oControl is TextBox)
                {
                    oControl.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                }
                if (oControl is MaskedTextBox)
                {
                    oControl.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                }
                if (oControl is Label)
                {
                    if (oControl.Name == "eti_tituloForm")
                    {
                        oControl.BackColor = ColorTranslator.FromHtml(fuerte); // cuando usamos FromHtml NO da error por fondo transparente
                    }
                    else
                    {
                        oControl.BackColor = ColorTranslator.FromHtml(normal); // cuando usamos FromHtml NO da error por fondo transparente
                    }
                }
                if (oControl is CheckBox)
                {
                    oControl.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                }
                if (oControl is RadioButton)
                {
                    oControl.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                }
                if (oControl is ListBox)
                {
                    oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                    oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                    oControl.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                }
                if (oControl is ComboBox)
                {
                    oControl.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                }
                if (oControl is Panel)
                {
                    oControl.BackColor = ColorTranslator.FromHtml(suave);
                    foreach (System.Windows.Forms.Control control in oControl.Controls)
                    {
                        if (control is TextBox)
                        {
                            control.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                        }
                        if (control is MaskedTextBox)
                        {
                            control.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                        }
                        if (control is Label)
                        {
                            control.BackColor = ColorTranslator.FromHtml(normal); // Color.FromArgb(1, 186, 218, 169);
                        }
                        if (control is CheckBox)
                        {
                            control.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                        }
                        if (control is RadioButton)
                        {
                            control.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                        }
                        if (control is ListBox)
                        {
                            control.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            control.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                            control.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                        }
                        if (control is ComboBox)
                        {
                            control.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                        }
                        if (control is Panel)
                        {
                            foreach (System.Windows.Forms.Control scontrol in control.Controls)
                            {
                                if (scontrol is TextBox)
                                {
                                    scontrol.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                                }
                                if (scontrol is MaskedTextBox)
                                {
                                    scontrol.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                                }
                                if (scontrol is Label)
                                {
                                    scontrol.BackColor = ColorTranslator.FromHtml(normal); // Color.FromArgb(1, 186, 218, 169);
                                }
                                if (scontrol is CheckBox)
                                {
                                    scontrol.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                                }
                                if (scontrol is RadioButton)
                                {
                                    scontrol.BackColor = ColorTranslator.FromHtml(suave); // Color.FromArgb(1, 186, 218, 169);
                                }
                                if (scontrol is ListBox)
                                {
                                    scontrol.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                                    scontrol.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                                    scontrol.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                                }
                                if (scontrol is ComboBox)
                                {
                                    scontrol.BackColor = ColorTranslator.FromHtml(suave); // cuando usamos FromHtml NO da error por fondo transparente
                                }
                            }
                        }
                    }
                }
            }
        }
        private void jalainfo()
        {
            // 31/07/2024 .. variabilizamos los datos que vamos a necesitar
            //nomForm = this.Name;
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
                tx_idOper.Text = Oegreso.IdMovim;
                tx_anno.Text = Oegreso.AnnoOp;    // DateTime.Parse(Oegreso.FechOper).Year.ToString();
                selecFecha1.Value = DateTime.Parse(Oegreso.FechOper);
            }
            else
            {
                tx_idOper.Text = "";
                if (Oegreso.FechOper == "") selecFecha1.Value = DateTime.Now.Date;
                else selecFecha1.Value = DateTime.Parse(Oegreso.FechOper);
                if (Oegreso.AnnoOp == "") tx_anno.Text = DateTime.Now.Date.Year.ToString();
                else tx_anno.Text = Oegreso.AnnoOp;
            }
            Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
            Tx_catEgre.Text = Oegreso.CatEgreso.largo; // Oegreso.CatEgreso.nombre
            eti_nomCat.Text = Oegreso.CatEgreso.nombre;  // Oegreso.CatEgreso.largo
            cmb_mon.SelectedValue = Oegreso.Moneda.codigo;
            tx_monto.Text = Oegreso.Monto.monOrige.ToString("#0.00");
            tx_tipcam.Text = Oegreso.TipCamb.ToString("#0.000");
            Tx_ctaDes.Text = Oegreso.CajaDes.largo;    // Oegreso.CajaDes.nombre
            eti_nomCaja.Text = Oegreso.CajaDes.nombre;   // Oegreso.CajaDes.largo
            Tx_nomProv.Text = Oegreso.Proveedor.nombre;
            tx_dat_provee.Text = Oegreso.Proveedor.codigo;
            tx_prov.Text = Oegreso.Proveedor.nombre;        // 29/11/2024 ahora estamos usando el autocompletado
            eti_nomprovee.Text = Oegreso.Proveedor.nombre;  // 31/08/2024 ya no es necesario
            tx_descrip.Text = Oegreso.Descrip;
            if (Ogiro != null && Ogiro.ctades != "")
            {
                chk_giroC.CheckState = CheckState.Checked;
                tx_ctaGiro.Text = Ogiro.largo;      // ctades;     // nombre corto
                tx_dat_giro.Text = Ogiro.idcod;     // idcodice de la cuenta
                eti_nomCtaGiro.Text = Ogiro.ctades; //largo;  // nombre largo de la cuenta
            }
            else
            {
                chk_giroC.CheckState = CheckState.Unchecked;
                tx_ctaGiro.Text = "";     // nombre corto
                tx_dat_giro.Text = "";     // idcodice de la cuenta
                eti_nomCtaGiro.Text = "";  // nombre largo de la cuenta
            }
        }                                                   // muestra en el formulario los objetos de la clase Egresos
        private void initCampos()
        {
            tx_anno.MaxLength = 4;
            Tx_catEgre.MaxLength = 50;  // (descrizionerid=20)(descrizione=50)
            Tx_catEgre.CharacterCasing = CharacterCasing.Upper;
            Tx_ctaDes.MaxLength = 50;  // (descrizionerid=20)(descrizione=50)
            Tx_ctaDes.CharacterCasing = CharacterCasing.Upper;
            tx_ctaGiro.MaxLength = 50;  // (descrizionerid=20)(descrizione=50)
            tx_ctaGiro.CharacterCasing = CharacterCasing.Upper;
            tx_descrip.MaxLength = 100;
            tx_idOper.MaxLength = 15;
            Tx_nomProv.MaxLength = 50;
            Tx_nomProv.CharacterCasing = CharacterCasing.Upper;
        }                                               // inicializa ancho de campos y upper case
        private void datsimil()
        {
            /*
            string fechUlt = ""; string fechAn = "";
            string[] ju = jala_ultimo(dt_grillaE, "EGRESO", ((rb_omg.Checked == true) ? "OMG" : "PER"), Tx_fecha.Text);
            //llenamos los objetos
            if (rb_omg.Checked == true && ju[3] != "")
            {
                Ocajd.largo = ju[8].ToString();     // ju[7].ToString()
                Ocajd.nombre = ju[7].ToString();    // ju[8].ToString()
                Ocajd.codigo = ju[6].ToString();
                OcatEg.largo = ju[3].ToString();    // ju[2].ToString()
                OcatEg.nombre = ju[2].ToString();   // ju[3].ToString()
                OcatEg.codigo = ju[1].ToString();
                Omone.nombre = ju[5].ToString();
                Omone.codigo = ju[4].ToString();
                Omone.siglas = ju[15].ToString();
                Omonto.tipCOri = decimal.Parse(ju[0]);
                Omonto.tipCDol = decimal.Parse(ju[0]);
                Omonto.codMOrige = ju[4].ToString();
                Omonto.monDolar = decimal.Parse(ju[17]);
                //Omonto.monEuros = ;
                Omonto.monSoles = decimal.Parse(ju[18]);
                Omonto.monOrige = decimal.Parse(ju[16]);
                Oprove.codigo = ju[9].ToString();
                Oprove.nombre = ju[10].ToString();
                Ogiro.largo = ju[14].ToString();
                Ogiro.idcod = ju[12].ToString();
                Ogiro.ctades = ju[13].ToString();
            }
            if (rb_pers.Checked == true && ju[3] != "")
            {
                // tipo de cambio,categoria id,categoria corto,categoria largo,moneda codigo,moneda nombre,cuenta id,cuenta corto
                //       0              1             2                3             4              5           6          7
                // cuenta largo,proveedor id,proveedor nombre,descripción,ctaGiro id,ctaGiro corto,ctaGiro largo,moneda siglas,
                //       8            9            10            11           12          13           14            15
                // monto original,monto en dolares,monto soles,fecha,año 
                //       16             17             18         19  20
                Ocajd.largo = ju[8].ToString();     // ju[7].ToString()
                Ocajd.nombre = ju[7].ToString();    // ju[8].ToString()
                Ocajd.codigo = ju[6].ToString();
                OcatEg.largo = ju[3].ToString();    // ju[2].ToString()
                OcatEg.nombre = ju[2].ToString();   // ju[3].ToString()
                OcatEg.codigo = ju[1].ToString();
                Omone.nombre = ju[5].ToString();
                Omone.codigo = ju[4].ToString();
                Omone.siglas = ju[15].ToString();
                Omonto.tipCOri = decimal.Parse(ju[0]);
                Omonto.monOrige = decimal.Parse(ju[16]);
                Omonto.monSoles = decimal.Parse(ju[18]);
                Omonto.monDolar = decimal.Parse(ju[17]);
                //Omonto.monEuros = 
                Omonto.monSoles = decimal.Parse(ju[18]);
                Omonto.monOrige = decimal.Parse(ju[16]);
                Oprove.codigo = ju[9].ToString();
                Oprove.nombre = ju[10].ToString();
                Ogiro.largo = ju[14].ToString();
                Ogiro.idcod = ju[12].ToString();
                Ogiro.ctades = ju[13].ToString();
            }
            fechAn = ju[20].ToString();
            fechUlt = ju[19].ToString();
            //
            Oegreso.creaEgreso(pan_p.Tag.ToString(), fechUlt, OcatEg, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                Ocajd, Oprove, ju[11].ToString(), "", Ogiro, fechAn);
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
                            tx_tipcam.Text = Math.Round(dr.GetDecimal(0), 3).ToString();
                            Omonto.tipCDol = Math.Round(dr.GetDecimal(0), 3);
                            //Omonto.tipCOri = Math.Round(dr.GetDecimal(1), 3);
                            if (Omonto.codMOrige != null && Omonto.codMOrige != "")
                            {
                                Omonto.tipCOri = (Omonto.codMOrige == codEur) ? Math.Round(dr.GetDecimal(1), 3) : (Omonto.codMOrige == codDol) ? Math.Round(dr.GetDecimal(0), 3) : Math.Round(dr.GetDecimal(0), 3);
                            }
                            tcDia.tcD = Omonto.tipCDol;
                            tcDia.tcE = Math.Round(dr.GetDecimal(1), 3);   // Omonto.tipCOri;
                            retorna = true;
                            if (Omonto.tipCDol <= 0 || tcDia.tcE <= 0) // Omonto.tipCDol <= 0 || Omonto.tipCOri <= 0
                            {
                                MessageBox.Show("El tipo de cambio Dólares es: " + Omonto.tipCDol.ToString() + Environment.NewLine +
                                    "El tipo de cambio Euros es: " + Omonto.tipCOri.ToString(), "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                retorna = false;
                                //this.Close();
                            }
                        }
                    }
                    else
                    {
                        var aa = MessageBox.Show("No existen tipos de cambio para la fecha actual" + Environment.NewLine +
                            "Debe ingresarlos en este momento", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        /*if (aa == DialogResult.Yes)
                        {
                            // llamada a formulario de tipos de cambio
                            tipcam f_tc = new tipcam();
                            f_tc.ShowDialog();
                        } */
                        this.Close();
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
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "NUEVO";
            if (tipCambio(null) == true)   // tx_tipcam.Text == ""
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
                chk_giroC.Enabled = true;
                selecFecha1.Value = DateTime.Now.Date;
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
            if (tipCambio(null) == true)
            {
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
                chk_giroC.Enabled = false;
                tx_ctaGiro.ReadOnly = true;
                pan_p.Enabled = true;
                rb_omg.Enabled = true;
                rb_pers.Enabled = true;
                tx_anno.Text = DateTime.Now.Year.ToString();
                tx_anno.ReadOnly = false;
                tx_idOper.ReadOnly = false;
                tx_idOper.Focus();
            }
            else
            {
                this.Close();
            }
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
            limpiaObj();
            limpiaTE();
            sololee("");
            chk_giroC.Enabled = false;
            pan_p.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
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
            limpiaObj();
            limpiaTE();
            sololee("");
            pan_p.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
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
            OcatEg.codigo = "";                                       // Objeto categoría de egreso
            OcatEg.nombre = "";
            OcatEg.largo = "";
            Omone.codigo = "";                                        // Objeto moneda
            Omone.nombre = "";
            Omone.siglas = "";
            Ocajd.codigo = "";                                        // Objeto cada de destino - desde donde sale el dinero
            Ocajd.nombre = "";
            Ocajd.largo = "";
            Oprove.codigo = "";                                       // Objeto proveedor
            Oprove.nombre = "";
            Omonto.codMOrige = "";                                    // Objeto monto
            Omonto.monDolar = 0;
            Omonto.monEuros = 0;
            Omonto.monOrige = 0;
            Omonto.monSoles = 0;
            Omonto.tipCDol = 0;
            Omonto.tipCOri = 0;
            Ogiro.ctades = "";
            Ogiro.tipodes = "";
            Ogiro.idcod = "";
            Ogiro.codigo = "";
            Oegreso.limpia();
        }
        private void limpiaTE() // limpia textbox, etiquetas, combos
        {
            tx_idOper.Clear();
            Tx_catEgre.Clear();
            Tx_ctaDes.Clear();
            tx_ctaGiro.Clear();
            tx_descrip.Clear();
            tx_monto.Clear();
            Tx_nomProv.Clear();
            tx_provee.Clear();
            tx_prov.Clear();
            tx_dat_provee.Clear();
            //tx_tipcam.Clear();    // 05-09-2024 mejor no limpiamos el tipo de cambio
            tx_dat_giro.Clear();
            selecFecha1.Value = DateTime.Now.Date;
            Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
            tx_anno.Text = DateTime.Now.Date.Year.ToString();
            //
            eti_nomCaja.Text = "";
            eti_nomCat.Text = "";
            eti_nomCtaGiro.Text = "";
            eti_nomprovee.Text = "";
            chk_datSimil.Checked = false;
            cmb_mon.SelectedIndex = -1; // no debe ser cero 02/09/2024 porque el objeto moneda esta limpio
            chk_giroC.Checked = false;
            // 
            Tx_ctaDes.BackColor = ColorTranslator.FromHtml(OColores.Fondo_suave);
        }
        private void escribe(string quien)  // pones los campos necesarios en readonly = false
        {
            if (quien == "EDICION") tx_idOper.ReadOnly = false;
            else tx_idOper.ReadOnly = true;
            Tx_catEgre.ReadOnly = false;
            Tx_ctaDes.ReadOnly = false;
            tx_ctaGiro.ReadOnly = false;
            tx_descrip.ReadOnly = false;
            tx_monto.ReadOnly = false;
            Tx_nomProv.ReadOnly = false;
            tx_provee.ReadOnly = true; // false; 31/08/2024 solo se jala con F1, no se puede validar por nombre
            tx_prov.ReadOnly = true;
            tx_tipcam.ReadOnly = false;
            //
            cmb_mon.Enabled = true;
            rb_omg.Enabled = true;
            rb_pers.Enabled = true;
            chk_datSimil.Enabled = true;
            chk_giroC.Enabled = true;
            cmb_mon.SelectedIndex = -1;
            cmb_mon_SelectedIndexChanged(null, null);
        }
        private void sololee(string quien)  //    // T=todos los campos, "" ó "C" campos comunes
        {
            Tx_catEgre.ReadOnly = true;
            Tx_ctaDes.ReadOnly = true;
            tx_ctaGiro.ReadOnly = true;
            tx_descrip.ReadOnly = true;
            tx_monto.ReadOnly = true;
            Tx_nomProv.ReadOnly = true;
            tx_provee.ReadOnly = true;
            tx_prov.ReadOnly = true;
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
                jalaGrilla(diasAtroya, "cassaomg");  // muestra datos de un dias atras hasta hoy
                Tx_ctaDes.Values = lista_DES.ToArray();
                tx_ctaGiro.Values = lista_CON.ToArray();   // lista_DES.ToArray();
            }
        }
        private void rb_pers_Click(object sender, EventArgs e)
        {
            if (rb_pers.Checked == true)
            {
                eti_tituloForm.Text = eti_tituloForm.Tag.ToString() + "DE CUENTAS PERSONALES";
                pan_p.Tag = "personal";
                limpiaTE();
                jalaGrilla(diasAtroya, "cassaconti");  // muestra datos de un dias atras hasta hoy
                Tx_ctaDes.Values = lista_CON.ToArray();
                tx_ctaGiro.Values = lista_CON.ToArray();
            }
        }
        private void chk_giroC_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_giroC.CheckState == CheckState.Checked)
            {
                tx_ctaGiro.Visible = true;
                eti_nomCtaGiro.Visible = true;
            }
            else
            {
                tx_ctaGiro.Visible = false;
                eti_nomCtaGiro.Visible = false;
                tx_ctaGiro.Text = "";
                eti_nomCtaGiro.Text = "";
                tx_dat_giro.Text = "";
                Ogiro.ctades = "";
                Ogiro.largo = "";
                Ogiro.idcod = "";
                Ogiro.tipodes = "";
                Ogiro.codigo = "";
            }
        }
        private void chk_datSimil_CheckStateChanged(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO")
            {
                if (chk_datSimil.CheckState == CheckState.Checked)
                {
                    if (Tx_catEgre.Text == "" && Tx_ctaDes.Text == "")
                    {
                        //if (advancedDataGridView1.Rows.Count > 0) datsimil(); deshabilitamos esto 30/01/2025,
                        // ahora datos similares es solo para el ultimo registro ingresado
                    }
                }
            }
        }

        #endregion

        #region enters, leaves y validaciones
        private void TextBox_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        }
        private void Tx_catEgre_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (Tx_catEgre.Text.Trim() != "")
                    {
                        //DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizionerid='" + Tx_catEgre.Text.Trim() + "'");
                        DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catEgre.Text.Trim() + "' and numero=1");
                        if (nc.Length > 0)
                        {
                            eti_nomCat.Text = nc[0].ItemArray[3].ToString();    // nc[0].ItemArray[2].ToString()
                            OcatEg.codigo = nc[0].ItemArray[1].ToString();
                            OcatEg.largo = Tx_catEgre.Text;    // OcatEg.nombre = Tx_catEgre.Text
                            OcatEg.nombre = eti_nomCat.Text;     // OcatEg.largo = eti_nomCat.Text
                        }
                        else
                        {
                            Tx_catEgre.Clear();
                            eti_nomCat.Text = "";
                            MessageBox.Show("No existe el nombre del egreso");
                        }
                    }
                }
            }
        }
        private void Tx_catEgre_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_catEgre.Text.Trim() != "")
                {
                    DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catEgre.Text.Trim() + "' and numero=1");
                    if (nc.Length > 0)
                    {
                        eti_nomCat.Text = nc[0].ItemArray[3].ToString();    // nc[0].ItemArray[2].ToString()
                        OcatEg.codigo = nc[0].ItemArray[1].ToString();
                        OcatEg.largo = Tx_catEgre.Text;    // OcatEg.nombre = Tx_catEgre.Text
                        OcatEg.nombre = eti_nomCat.Text;     // OcatEg.largo = eti_nomCat.Text
                    }
                    else
                    {
                        Tx_catEgre.Clear();
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
                    xxx();  // esto no se ejecuta nunca!
                }
            }
        }
        private void Tx_ctaDes_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_ctaDes.Text.Trim() != "" && Tx_ctaDes.Text.Length >= 3)
                {
                    string[] vuelto = ValiCtaCon(Tx_ctaDes.Text, (rb_omg.Checked == true) ? "OMG" : "PER", "algo");
                    if (vuelto.Length > 0 && vuelto[0] != "")
                    {
                        Ocajd.codigo = vuelto[0];
                        Ocajd.nombre = vuelto[1];
                        Ocajd.largo = vuelto[2];
                        eti_nomCaja.Text = Ocajd.nombre; //  = Ocajd.largo
                        //
                        using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                        {
                            try
                            {
                                conn.Open();
                            }
                            catch (MySqlException ex)
                            {
                                MessageBox.Show(ex.Message, "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                Application.Exit();
                                return;
                            }
                            if (conn.State == ConnectionState.Open)
                            {
                                string nt = (rb_omg.Checked == true) ? "cassaomg" : "cassaconti";
                                using (MySqlCommand mic = new MySqlCommand("reps_saldoIniS", conn))  // reps_saldoIni
                                {
                                    mic.CommandType = CommandType.StoredProcedure;  // retorna valor en dólares
                                    mic.Parameters.AddWithValue("v_tabla", nt);
                                    mic.Parameters.AddWithValue("v_a01", "LIM");
                                    mic.Parameters.AddWithValue("v_a02", "");
                                    mic.Parameters.AddWithValue("v_a03", "");
                                    mic.Parameters.AddWithValue("v_a02f", selecFecha1.Value.Date.AddDays(1).ToString("yyyy-MM-dd"));
                                    mic.Parameters.AddWithValue("v_a04", "");
                                    mic.Parameters.AddWithValue("v_a05", Ocajd.codigo);
                                    using (MySqlDataReader dr = mic.ExecuteReader())
                                    {
                                        if (dr.HasRows)
                                        {
                                            if (dr.Read())
                                            {
                                                if (!String.IsNullOrEmpty(dr[0].ToString()))
                                                {
                                                    double valor = 0;
                                                    //if (String.IsNullOrEmpty(Omonto.monDolar.ToString())) valor = dr.GetDouble(2);
                                                    //else valor = dr.GetDouble(2) - (double)Omonto.monDolar;
                                                    valor = dr.GetDouble(2);        // saldo final en soles
                                                    if (Omone.codigo == "MON001") valor = valor - (double)Omonto.monOrige;
                                                    if (Omone.codigo == "MON002") valor = valor - (double)(Omonto.monOrige * Omonto.tipCOri);
                                                    if (Omone.codigo == "MON003") valor = valor - (double)(Omonto.monOrige * Omonto.tipCOri);
                                                    if (valor <= 0) Tx_ctaDes.ForeColor = System.Drawing.Color.DarkRed;
                                                    else Tx_ctaDes.ForeColor = System.Drawing.Color.Black;
                                                }
                                                else MessageBox.Show("No existen datos de" + Environment.NewLine +
                                                    "saldo en la cuenta", "Valores nulos");
                                                tx_provee.Focus();
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("No se encuentra datos de la cuenta" + Environment.NewLine +
                                                "para calcular su saldo", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Tx_ctaDes.Clear();
                        eti_nomCaja.Text = "";
                        MessageBox.Show("No existe el nombre de la cuenta");
                    }
                }
            }
        }
        private void Tx_provee_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (tx_prov.Text.Trim() != "")    // tx_provee.Text.Trim() != ""
                {
                    Oprove = ValiProvee(tx_dat_provee.Text);
                    if (Oprove.nombre == "")
                    {
                        eti_nomprovee.Text = "";
                        //tx_provee.Text = "";
                        tx_prov.Text = "";
                        MessageBox.Show("No existe el código de proveedor");
                    }
                    else
                    {
                        //tx_provee.Text = Oprove.nombre;
                        tx_prov.Text = Oprove.nombre;
                        eti_nomprovee.Text = Oprove.nombre; // 31/08/2024 ya no lo usamos
                    }
                }
                else
                {
                    eti_nomprovee.Text = "";
                    //tx_provee.Text = "";
                    tx_prov.Text = "";
                }
            }
        }   // 29/11/2024 ya no, ahora usamos el autocompletado tx_nomprov.text
        private void tx_idOper_Validating(object sender, CancelEventArgs e)       // busca en toda la base de datos
        {
            if (tx_idOper.Text.Trim() != "" && !("NUEVO").Contains(Tx_modo.Text))   //   tx_idOper.Text.Trim() != "" && ("NUEVO,EDICION").Contains(Tx_modo.Text)
            {
                string[] retu = ValiIdOper((rb_omg.Checked == true) ? "OMG" : "PER", tx_idOper.Text.Trim(), tx_anno.Text, "S");
                if (retu[0] == "")
                {
                    limpiaObj();
                    limpiaTE();
                    MessageBox.Show("No existe el código de operación");
                }
                else
                {
                    // asignamos los valores de retu[] a los objetos
                    string fecOp = "";              // fecha de operacion
                    decimal tipca = 0;              // tip cambio del monto origen
                    string descr = "";              // descripcion de la operacion
                    string idmov = "";              // id del movimiento
                    if (rb_omg.Checked == true)
                    {
                        fecOp = retu[2].Substring(0, 10);       // fecha
                        OcatEg.codigo = retu[19];               // IDCategoria
                        OcatEg.nombre = retu[4];                // EGRESO
                        OcatEg.largo = retu[24];                // "DET_EGRESO"
                        Omone.codigo = retu[20];                // "codimon"
                        Omone.siglas = retu[5];                 // "MONEDA"
                        Omone.nombre = retu[21];                // "nombmon"
                        Omonto.codMOrige = retu[20];            // "codimon"
                        Omonto.monOrige = decimal.Parse(retu[6]);   // "MONTO"
                        Omonto.tipCOri = decimal.Parse(retu[22]);   // "TCMonOri" 
                        Omonto.monDolar = decimal.Parse(retu[15]);  // "ImportoDU"
                        Omonto.tipCDol = decimal.Parse(retu[8]);    // "TIP_CAMBIO"
                        Omonto.monSoles = decimal.Parse(retu[16]);  // "ImportoSU"
                        tipca = decimal.Parse(retu[22]);            // "TCMonOri"
                        Ocajd.codigo = retu[18];                // "IDDestino"
                        Ocajd.nombre = retu[3];                 // "DESTINO"
                        Ocajd.largo = retu[23];                 // "DET_DESTINO"
                        Oprove.codigo = retu[17];               // "idanagrafica"
                        Oprove.nombre = retu[9];                // "PROVEEDOR"
                        descr = retu[7];                        // "DESCRIPCION"
                        idmov = retu[1];                        // "ID_MOVIM"
                        Ogiro.ctades = retu[26];                // nombre corto de la cuenta CTA_GIRO
                        Ogiro.tipodes = retu[10];               // tipodesgiro
                        Ogiro.codigo = retu[25];                // identificador del giroconto, campo CodGiro 
                        Ogiro.idcod = retu[11];                 // idcodice de la cuenta idgiroconto
                        Ogiro.largo = retu[12];                 // nombre largo de la cuenta CTA_DESTINO
                    }
                    else
                    {
                        fecOp = retu[2].Substring(0, 10);       // "FECHA"
                        OcatEg.codigo = retu[19];               // "IDCategoria"
                        OcatEg.nombre = retu[4];                // "EGRESO"
                        OcatEg.largo = retu[24];                // "DET_EGRESO"
                        Omone.codigo = retu[20];                // "codimon"
                        Omone.siglas = retu[5];                 // "MONEDA"
                        Omone.nombre = retu[21];                // "nombmon"
                        Omonto.codMOrige = retu[20];            // "codimon"
                        Omonto.monOrige = decimal.Parse(retu[6]);   // "MONTO"
                        Omonto.tipCOri = decimal.Parse(retu[22]);   // "TCMonOri"
                        Omonto.monDolar = decimal.Parse(retu[15]);  // "ImportoDU"
                        Omonto.tipCDol = decimal.Parse(retu[8]);    // "TIP_CAMBIO"
                        Omonto.monSoles = decimal.Parse(retu[16]);  // "ImportoSU"
                        tipca = decimal.Parse(retu[22]);        // "TCMonOri"
                        Ocajd.codigo = retu[18];                // "IDConto"
                        Ocajd.nombre = retu[3];                 // "CUENTA"
                        Ocajd.largo = retu[23];                 // "DET_CUENTA"
                        Oprove.codigo = retu[17];               // "idanagrafica"
                        Oprove.nombre = retu[9];                // "PROVEEDOR"
                        descr = retu[7];                        // "DESCRIPCION"
                        idmov = retu[1];                        // "ID_MOVIM"
                        Ogiro.ctades = retu[26];                // nombre corto de la cuenta
                        Ogiro.tipodes = retu[10];               // OMG-PER
                        Ogiro.codigo = retu[25];                // identificador del giroconto, campo CodGiro 
                        Ogiro.idcod = retu[11];                 // idcodice de la cuenta
                        Ogiro.largo = retu[12];                 // nombre largo de la cuenta CTA_DESTINO
                    }
                    Oegreso.creaEgreso(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, tipca,
                            Ocajd, Oprove, descr, idmov, Ogiro, tx_anno.Text);
                    jalaoc();
                }
            }
        }
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
                    Omonto.tipCDol = tcDia.tcD; // Omonto.tipCOri;
                    Omonto.tipCOri = tcDia.tcD;
                    Omonto.monEuros = 0;
                    Omonto.monDolar = decimal.Parse(tx_monto.Text);
                    Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                    calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), decimal.Parse(tx_tipcam.Text));
                }
                if (Omone.codigo == codSol)
                {
                    Omonto.tipCDol = tcDia.tcD;
                    Omonto.tipCOri = tcDia.tcD;
                    Omonto.monEuros = 0;
                    Omonto.monSoles = decimal.Parse(tx_monto.Text); // Omonto.monDolar * Omonto.tipCOri;
                    Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                    calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCDol);
                }
                if (Omone.codigo == codEur)
                {
                    Omonto.tipCDol = 0;
                    Omonto.tipCOri = tcDia.tcE;
                    Omonto.monEuros = decimal.Parse(tx_monto.Text);
                    Omonto.monDolar = 0;
                    Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                    calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCOri);
                }
            }
        }
        private void tx_tipcam_Validating(object sender, CancelEventArgs e)
        {
            decimal monti = 0; decimal cambi = 0;
            decimal.TryParse(tx_monto.Text, out monti);
            decimal.TryParse(tx_tipcam.Text, out cambi);
            tx_tipcam.Text = Math.Round(cambi, 3).ToString("#0.000");
            if (Tx_modo.Text == "NUEVO" && monti > 0)
            {
                Omonto.monOrige = monti;
                if (true)
                {
                    //calc_monedas(cmb_mon, monti, cambi);
                    if (Omonto.codMOrige == codDol) Omonto = calc_monedas(cmb_mon, monti, cambi);
                    if (Omonto.codMOrige == codSol) Omonto = calc_monedas(cmb_mon, monti, cambi);
                    if (Omonto.codMOrige == codEur) Omonto = calc_monedas(cmb_mon, monti, Omonto.tipCOri);
                }
            }
        }
        public string[] ValiIdOper(string _tipo_, string _id_, string yea, string _EoS_) // EoS => E entrada | S salida
        {
            string[] retorna = { "", "", "", "", "", "", "", "", "", "",
                                "", "", "", "", "", "", "", "", "", "", 
                                "", "", "", "", "", "", "", "", "", "",
                                "", ""};           // 32 valores
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        string consulta = "";
                        if (_tipo_.ToUpper() == "PRE-OMG")
                        {
                            consulta = "ConPrel_cassaOmg";
                        }
                        if (_tipo_.ToUpper() == "PRE-PER")
                        {
                            consulta = "ConPrel_cassaConti";
                        }
                        if (_tipo_.ToUpper() == "OMG")
                        {
                            if (_EoS_ == "E") consulta = "ConIngre_cassaOmg";
                            if (_EoS_ == "S") consulta = "ConEgre_cassaOmg";
                        }
                        if (_tipo_.ToUpper() == "PER")
                        {
                            if (_EoS_ == "E") consulta = "ConIngre_cassaConti";
                            if (_EoS_ == "S") consulta = "ConEgre_cassaConti";
                        }
                        if (_tipo_.ToUpper() == "CAMION")
                        {
                            consulta = "ConCamion";
                        }
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            _id_ = CDerecha("0000000000000" + _id_, 15);
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@Vdias", 0);
                            micon.Parameters.AddWithValue("@Vanno", int.Parse(yea));
                            micon.Parameters.AddWithValue("@Vidmov", _id_);
                            if (_tipo_ == "PRE-OMG" || _tipo_ == "PRE-PER") micon.Parameters.AddWithValue("@proce", "T");
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.HasRows == true)
                                {
                                    if (dr.Read())
                                    {
                                        if (_tipo_.ToUpper() == "PRE-OMG" || _tipo_.ToUpper() == "PRE-PER")
                                        {
                                            if (dr[0] != null && dr[0].ToString() != "")
                                            {
                                                if (_tipo_.ToUpper() == "PRE-OMG")
                                                {
                                                    // ANNO,ID_MOVIM,FECHA,DET_CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,OPERADOR,
                                                    // dia,APROBADOR,FEC_PROCESO,GIRO_CTA,IDGiroConto,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,
                                                    // codimon,nombmon,TCMonOri,CUENTA,DET_EGRESO,CodGiro,CTA_DESTINO,CTA_GIRO,RUC,cuentaB,CASA,tipoE
                                                    retorna[0] = dr["ANNO"].ToString();
                                                    retorna[1] = dr["ID_MOVIM"].ToString();
                                                    retorna[2] = dr["FECHA"].ToString();
                                                    retorna[3] = dr["DET_CUENTA"].ToString();
                                                    retorna[4] = dr["EGRESO"].ToString();
                                                    retorna[5] = dr["MONEDA"].ToString();
                                                    retorna[6] = dr["MONTO"].ToString();
                                                    retorna[7] = dr["DESCRIPCION"].ToString();
                                                    retorna[8] = dr["TIP_CAMBIO"].ToString();
                                                    retorna[9] = dr["PROVEEDOR"].ToString();
                                                    retorna[10] = dr["OPERADOR"].ToString();
                                                    retorna[11] = dr["dia"].ToString();
                                                    retorna[12] = dr["APROBADOR"].ToString();
                                                    retorna[13] = dr["FEC_PROCESO"].ToString();
                                                    retorna[14] = dr["GIRO_CTA"].ToString();
                                                    retorna[15] = dr["IDGiroConto"].ToString();
                                                    retorna[16] = dr["ImportoDU"].ToString();
                                                    retorna[17] = dr["ImportoSU"].ToString();
                                                    retorna[18] = dr["idanagrafica"].ToString();
                                                    retorna[19] = dr["IDConto"].ToString();
                                                    retorna[20] = dr["IDCategoria"].ToString();
                                                    retorna[21] = dr["codimon"].ToString();
                                                    retorna[22] = dr["nombmon"].ToString();
                                                    retorna[23] = dr["TCMonOri"].ToString();
                                                    retorna[24] = dr["CUENTA"].ToString();
                                                    retorna[25] = dr["DET_EGRESO"].ToString();
                                                    retorna[26] = dr["CodGiro"].ToString();
                                                    retorna[27] = dr["CTA_DESTINO"].ToString();
                                                    retorna[28] = dr["CTA_GIRO"].ToString();
                                                    retorna[29] = dr["CASA"].ToString();
                                                    retorna[30] = dr["tipoE"].ToString();
                                                    retorna[31] = dr["RUC"].ToString();             // ruc del proveedor
                                                    retorna[32] = dr["cuentaB"].ToString();          // cuenta bancaria del proveedor
                                                }
                                                if (_tipo_.ToUpper() == "PRE-PER")
                                                {
                                                    // ANNO,ID_MOVIM,FECHA,DET_CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,OPERADOR,
                                                    // dia,APROBADOR,FEC_PROCESO,GIRO_CTA,IDGiroConto,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,
                                                    // codimon,nombmon,TCMonOri,CUENTA,DET_EGRESO,CodGiro,CTA_DESTINO,CTA_GIRO,RUC,cuentaB,CASA,tipoE
                                                    retorna[0] = dr["ANNO"].ToString();
                                                    retorna[1] = dr["ID_MOVIM"].ToString();
                                                    retorna[2] = dr["FECHA"].ToString();
                                                    retorna[3] = dr["DET_CUENTA"].ToString();
                                                    retorna[4] = dr["EGRESO"].ToString();
                                                    retorna[5] = dr["MONEDA"].ToString();
                                                    retorna[6] = dr["MONTO"].ToString();
                                                    retorna[7] = dr["DESCRIPCION"].ToString();
                                                    retorna[8] = dr["TIP_CAMBIO"].ToString();
                                                    retorna[9] = dr["PROVEEDOR"].ToString();
                                                    retorna[10] = dr["OPERADOR"].ToString();
                                                    retorna[11] = dr["dia"].ToString();
                                                    retorna[12] = dr["APROBADOR"].ToString();
                                                    retorna[13] = dr["FEC_PROCESO"].ToString();
                                                    retorna[14] = dr["GIRO_CTA"].ToString();
                                                    retorna[15] = dr["IDGiroConto"].ToString();
                                                    retorna[16] = dr["ImportoDU"].ToString();
                                                    retorna[17] = dr["ImportoSU"].ToString();
                                                    retorna[18] = dr["idanagrafica"].ToString();
                                                    retorna[19] = dr["IDConto"].ToString();
                                                    retorna[20] = dr["IDCategoria"].ToString();
                                                    retorna[21] = dr["codimon"].ToString();
                                                    retorna[22] = dr["nombmon"].ToString();
                                                    retorna[23] = dr["TCMonOri"].ToString();
                                                    retorna[24] = dr["CUENTA"].ToString();
                                                    retorna[25] = dr["DET_EGRESO"].ToString();
                                                    retorna[26] = dr["CodGiro"].ToString();
                                                    retorna[27] = dr["CTA_DESTINO"].ToString();
                                                    retorna[28] = dr["CTA_GIRO"].ToString();
                                                    retorna[29] = dr["CASA"].ToString();
                                                    retorna[30] = dr["tipoE"].ToString();
                                                    retorna[31] = dr["RUC"].ToString();             // ruc del proveedor
                                                    retorna[32] = dr["cuentaB"].ToString();          // cuenta bancaria del proveedor
                                                }
                                            }
                                        }
                                        if (_tipo_.ToUpper() == "PER" || _tipo_.ToUpper() == "OMG")
                                        {
                                            if (dr[0] != null && dr[0].ToString() != "")
                                            {
                                                if (_tipo_.ToUpper() == "PER" && _EoS_ == "S")
                                                {
                                                    // ANNO,ID_MOVIM,FECHA,CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,IDGiroConto,CTA_DESTINO,
                                                    //   0      1      2      3      4     5      6        7          8          9        10         11          12
                                                    // usuario,dia,ImportoDU,ImportoSU,a.idanagrafica,a.IDConto,a.IDCategoria,codimon,a.nombmon,a.TCMonOri,DET_CUENTA,
                                                    //    13    14     15         16          17           18         19         20       21         22         23
                                                    // DET_EGRESO, a.CodGiro,CASA,CTA_GIRO
                                                    //      24          25    26     27
                                                    retorna[0] = dr["ANNO"].ToString();
                                                    retorna[1] = dr["ID_MOVIM"].ToString();
                                                    retorna[2] = dr["FECHA"].ToString();
                                                    retorna[3] = dr["CUENTA"].ToString();
                                                    retorna[4] = dr["EGRESO"].ToString();
                                                    retorna[5] = dr["MONEDA"].ToString();
                                                    retorna[6] = dr["MONTO"].ToString();
                                                    retorna[7] = dr["DESCRIPCION"].ToString();
                                                    retorna[8] = dr["TIP_CAMBIO"].ToString();
                                                    retorna[9] = dr["PROVEEDOR"].ToString();
                                                    retorna[10] = dr["GIRO_CTA"].ToString();        // tipo cta destino OMG o PER
                                                    retorna[11] = dr["IDGiroConto"].ToString();     // cuenta destino del giro
                                                    retorna[12] = dr["CTA_DESTINO"].ToString();     // nombre largo de la cuenta destno
                                                    retorna[13] = dr["usuario"].ToString();
                                                    retorna[14] = dr["dia"].ToString();
                                                    retorna[15] = dr["ImportoDU"].ToString();
                                                    retorna[16] = dr["ImportoSU"].ToString();
                                                    retorna[17] = dr["idanagrafica"].ToString();
                                                    retorna[18] = dr["IDConto"].ToString();
                                                    retorna[19] = dr["IDCategoria"].ToString();
                                                    retorna[20] = dr["codimon"].ToString();
                                                    retorna[21] = dr["nombmon"].ToString();
                                                    retorna[22] = dr["TCMonOri"].ToString();
                                                    retorna[23] = dr["DET_CUENTA"].ToString();
                                                    retorna[24] = dr["DET_EGRESO"].ToString();
                                                    retorna[25] = dr["CodGiro"].ToString();
                                                    retorna[26] = dr["CTA_GIRO"].ToString(); // nombre corto de la cuenta giro
                                                }
                                                if (_tipo_.ToUpper() == "OMG" && _EoS_ == "S")
                                                {
                                                    // ANNO,ID_MOVIM,FECHA,DESTINO,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,idgiroconto,
                                                    // CTA_DESTINO,usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDDestino,IDCategoria,codimon,nombmon,TCMonOri,
                                                    // DET_DESTINO,DET_EGRESO,tipodesgiro,CodGiro,CTA_GIRO,CASA
                                                    retorna[0] = dr["ANNO"].ToString();
                                                    retorna[1] = dr["ID_MOVIM"].ToString();
                                                    retorna[2] = dr["FECHA"].ToString();
                                                    retorna[3] = dr["DESTINO"].ToString();
                                                    retorna[4] = dr["EGRESO"].ToString();
                                                    retorna[5] = dr["MONEDA"].ToString();
                                                    retorna[6] = dr["MONTO"].ToString();
                                                    retorna[7] = dr["DESCRIPCION"].ToString();
                                                    retorna[8] = dr["TIP_CAMBIO"].ToString();
                                                    retorna[9] = dr["PROVEEDOR"].ToString();
                                                    retorna[10] = dr["GIRO_CTA"].ToString();        // tipo cta destino OMG o PER
                                                    retorna[11] = dr["IDGiroConto"].ToString();     // cuenta destino del giro
                                                    retorna[12] = dr["CTA_DESTINO"].ToString();     // nombre largo de la cuenta destno
                                                    retorna[13] = dr["usuario"].ToString();
                                                    retorna[14] = dr["dia"].ToString();
                                                    retorna[15] = dr["ImportoDU"].ToString();
                                                    retorna[16] = dr["ImportoSU"].ToString();
                                                    retorna[17] = dr["idanagrafica"].ToString();
                                                    retorna[18] = dr["IDDestino"].ToString();
                                                    retorna[19] = dr["IDCategoria"].ToString();
                                                    retorna[20] = dr["codimon"].ToString();
                                                    retorna[21] = dr["nombmon"].ToString();
                                                    retorna[22] = dr["TCMonOri"].ToString();
                                                    retorna[23] = dr["DET_DESTINO"].ToString();
                                                    retorna[24] = dr["DET_EGRESO"].ToString();
                                                    retorna[25] = dr["CodGiro"].ToString();
                                                    retorna[26] = dr["CTA_GIRO"].ToString(); // nombre corto de la cuenta giro
                                                }
                                                if (_tipo_.ToUpper() == "PER" && _EoS_ == "E")
                                                {
                                                    // ANNO,ID_MOVIM,FECHA,CUENTA,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                                                    // IDConto,a.IDCategoria,a.codimon,a.nombmon,a.TCMonOri,DET_CUENTA,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                                                    // CASA,GIRO_CTA,a.IDGiroConto,CTA_DESTINO
                                                    retorna[0] = dr["ANNO"].ToString();
                                                    retorna[1] = dr["ID_MOVIM"].ToString();
                                                    retorna[2] = dr["FECHA"].ToString();
                                                    retorna[3] = dr["CUENTA"].ToString();
                                                    retorna[4] = dr["INGRESO"].ToString();
                                                    retorna[5] = dr["MONEDA"].ToString();
                                                    retorna[6] = dr["MONTO"].ToString();
                                                    retorna[7] = dr["DESCRIPCION"].ToString();
                                                    retorna[8] = dr["TIP_CAMBIO"].ToString();
                                                    retorna[9] = dr["usuario"].ToString();
                                                    retorna[10] = dr["dia"].ToString();
                                                    retorna[11] = dr["ImportoDE"].ToString();        // tipo cta destino OMG o PER
                                                    retorna[12] = dr["ImportoSE"].ToString();     // cuenta destino del giro
                                                    retorna[13] = dr["IDConto"].ToString();     // nombre largo de la cuenta destno
                                                    retorna[14] = dr["IDCategoria"].ToString();
                                                    retorna[15] = dr["codimon"].ToString();
                                                    retorna[16] = dr["nombmon"].ToString();
                                                    retorna[17] = dr["TCMonOri"].ToString();
                                                    retorna[18] = dr["DET_CUENTA"].ToString();
                                                    retorna[19] = dr["DET_INGRESO"].ToString();
                                                    retorna[20] = "";   // tipodesgiro
                                                    retorna[21] = "";   // CodGiro
                                                    retorna[22] = "";   // CTA_GIRO
                                                    retorna[23] = dr["CASA"].ToString();
                                                    retorna[24] = "";   // GIRO_CTA
                                                    retorna[25] = "";   // IDGiroConto
                                                    retorna[26] = dr["CTA_DESTINO"].ToString(); // nombre largo de la cuenta giro
                                                }
                                                if (_tipo_.ToUpper() == "OMG" && _EoS_ == "E")
                                                {
                                                    // ANNO,ID_MOVIM,FECHA,DESTINO,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                                                    // IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                                                    // CASA,GIRO_CTA,a.idgiroconto,CTA_DESTINO
                                                    retorna[0] = dr["ANNO"].ToString();
                                                    retorna[1] = dr["ID_MOVIM"].ToString();
                                                    retorna[2] = dr["FECHA"].ToString();
                                                    retorna[3] = dr["DESTINO"].ToString();
                                                    retorna[4] = dr["INGRESO"].ToString();
                                                    retorna[5] = dr["MONEDA"].ToString();
                                                    retorna[6] = dr["MONTO"].ToString();
                                                    retorna[7] = dr["DESCRIPCION"].ToString();
                                                    retorna[8] = dr["TIP_CAMBIO"].ToString();
                                                    retorna[9] = dr["usuario"].ToString();
                                                    retorna[10] = dr["dia"].ToString();
                                                    retorna[11] = dr["ImportoDE"].ToString();        // tipo cta destino OMG o PER
                                                    retorna[12] = dr["ImportoSE"].ToString();     // cuenta destino del giro
                                                    retorna[13] = dr["IDDestino"].ToString();     // nombre largo de la cuenta destno
                                                    retorna[14] = dr["IDCategoria"].ToString();
                                                    retorna[15] = dr["codimon"].ToString();
                                                    retorna[16] = dr["nombmon"].ToString();
                                                    retorna[17] = dr["TCMonOri"].ToString();
                                                    retorna[18] = dr["DET_DESTINO"].ToString();
                                                    retorna[19] = dr["DET_INGRESO"].ToString();
                                                    retorna[20] = "";   // tipodesgiro
                                                    retorna[21] = "";   // CodGiro 
                                                    retorna[22] = "";   // CTA_GIRO
                                                    retorna[23] = dr["CASA"].ToString();
                                                    retorna[24] = "";   // GIRO_CTA
                                                    retorna[25] = "";   // idgiroconto
                                                    retorna[26] = dr["CTA_DESTINO"].ToString();
                                                }
                                            }
                                        }
                                        if (_tipo_.ToUpper() == "CAMION")
                                        {
                                            if (dr[0] != null && dr[0].ToString() != "")
                                            {
                                                // ANNO,ID_MOVIM,FECHA,CAMION,DESTINO,TOTAL_SOL,DESCRIPCION,usuario,CASA,
                                                //   0     1       2      3      4       5           6         7      8  
                                                // TOTAL_DOL,TIP_CAMBIO,dia,codimon,nombmon,MONEDA,IDDestino,DET_DESTINO,
                                                //     9         10      11    12      13     14      15         16   
                                                // ImpCarbS,ImpViaS,ImpRicS,ImpVariS,imphons,ImpImpS,repdol,IDCamion
                                                //    17      18      19       20       21      22     23      24
                                                retorna[0] = dr["ANNO"].ToString();
                                                retorna[1] = dr["ID_MOVIM"].ToString();
                                                retorna[2] = dr["FECHA"].ToString();
                                                retorna[3] = dr["CAMION"].ToString();
                                                retorna[4] = dr["DESTINO"].ToString();
                                                retorna[5] = dr["TOTAL_SOL"].ToString();
                                                retorna[6] = dr["DESCRIPCION"].ToString();
                                                retorna[7] = dr["usuario"].ToString();
                                                retorna[8] = dr["CASA"].ToString();
                                                retorna[9] = dr["TOTAL_DOL"].ToString();
                                                retorna[10] = dr["TIP_CAMBIO"].ToString();
                                                retorna[12] = dr["codimon"].ToString();
                                                retorna[13] = dr["nombmon"].ToString();
                                                retorna[14] = dr["MONEDA"].ToString();
                                                retorna[15] = dr["IDDestino"].ToString();
                                                retorna[16] = dr["DET_DESTINO"].ToString();
                                                retorna[17] = dr["ImpCarbS"].ToString();
                                                retorna[18] = dr["ImpViaS"].ToString();
                                                retorna[19] = dr["ImpRicS"].ToString();
                                                retorna[20] = dr["ImpVariS"].ToString();
                                                retorna[21] = dr["imphons"].ToString();
                                                retorna[22] = dr["ImpImpS"].ToString();
                                                retorna[23] = dr["repdol"].ToString();
                                                retorna[24] = dr["IDCamion"].ToString();
                                                retorna[25] = "";
                                                retorna[26] = "";
                                                //retorna[27] = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error de conexión al servidor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // limpiamos ???  no, no, limpiamos en el llamador de la funcion
                }
            }
            return retorna;
        }                                           // valida idOper, si hay jala datos, sino No
        public provees ValiProvee(string idAnag)
        {
            provees retona = new provees();
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        using (MySqlCommand micon = new MySqlCommand("select ragionesociale from anag_for where TRIM(idanagrafica)=@codi", conn))
                        {
                            micon.Parameters.AddWithValue("@codi", idAnag.Trim());
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.HasRows == true)
                                {
                                    if (dr.Read())
                                    {
                                        if (dr[0] != null && dr[0].ToString() != "")
                                        {
                                            retona.codigo = idAnag;
                                            retona.nombre = dr[0].ToString();
                                        }
                                    }
                                }
                                else
                                {
                                    retona.codigo = "";
                                    retona.nombre = "";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error de conexión al servidor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    retona.codigo = "";
                    retona.nombre = "";
                }
            }
            return retona;
        }                               // valida existencia del proveedor
        public bool ValiCtaCon(string _nombre, string _tipo)    // OMG | PER
        {
            // validamos la existencia del nombre en ... descrizione
            bool retorna = false;
            DataRow[] row;  // Program.dt_definic.Select("idtabella='DES' and descrizionerid='" + _nombre + "'");
            if (_tipo == "OMG") row = Program.dt_definic.Select("idtabella='DES' and descrizione='" + _nombre + "'");
            else row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + _nombre + "'");
            foreach (DataRow dat in row)    // "idtabella='CON' and descrizionerid='" + _nombre + "'")
            {
                retorna = true;
            }
            return retorna;
        }                                 // valida existencia de la cuenta destino
        public string[] ValiCtaCon(string _nombre, string _tipo, string algo)    // OMG | PER
        {
            // validamos la existencia del nombre en ... descrizionerid
            string[] retorna = { "", "", ""};
            DataRow[] row;  // Program.dt_definic.Select("idtabella='DES' and descrizionerid='" + _nombre + "'")
            if (_tipo == "OMG") row = Program.dt_definic.Select("idtabella='DES' and descrizione='" + _nombre + "'");
            else row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + _nombre + "'");
            foreach (DataRow dat in row)    // "idtabella='CON' and descrizionerid='" + _nombre + "'")
            {
                retorna[0] = dat[1].ToString();   // codigo
                retorna[1] = dat[3].ToString().ToUpper();    // corto
                retorna[2] = dat[2].ToString().ToUpper();    // largo
            }
            return retorna;
        }
        public bool Vali_CAM(string _nombre)
        {
            // validamos la existencia del nombre en ... descrizionerid
            bool retorna = false;
            DataRow[] row = Program.dt_definic.Select("descrizionerid='" + _nombre + "' and idtabella='CAM'");
            foreach (DataRow dat in row)
            {
                retorna = true;
            }
            return retorna;
        }
        private void tx_tipcam_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                tx_idOper.Focus();
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
        private void selecFecha1_ValueChanged(object sender, EventArgs e)
        {
            // En ningun caso la fecha puede ser posterior al actual
            // si es nuevo la fecha puede ser anterior
            // si es edicion no se permite cambiar la fecha, 04/12/2024
            if (selecFecha1.Value.Date > DateTime.Now.Date)
            {
                MessageBox.Show("No se permite fechas posteriores", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                selecFecha1.Value = DateTime.Now.Date;
                Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
            }
            try
            {
                if ((Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION") && chk_datSimil.Checked == false)
                {
                    Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
                    tipCambio(null);
                }
            }
            catch (Exception ex)
            {
                Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
                tipCambio(null);
            }
        }
        private void tx_ctaGiro_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (tx_ctaGiro.Text.Trim() != "")
                    {
                        yyy();
                    }
                }
            }
        }
        private void Tx_nomProv_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_nomProv.Text.Trim().Length>3 && (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION"))
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (Tx_nomProv.Text.Trim() != "")
                    {
                        DataRow[] row;
                        {
                            row = dtpro.Select("nombre='" + Tx_nomProv.Text.Trim() + "'");
                        }
                        if (row.Length > 0)
                        {
                            tx_dat_provee.Text = row[0][0].ToString();     // idanagrafica
                            tx_prov.Text = Tx_nomProv.Text.Trim();
                            eti_nomprovee.Text = Tx_nomProv.Text.Trim(); // 31/08/2024 ya no usamos
                            Oprove.nombre = tx_prov.Text;
                            Oprove.codigo = tx_dat_provee.Text;
                        }
                        else
                        {
                            Tx_nomProv.Clear();
                            tx_dat_provee.Clear();
                            tx_prov.Clear();
                            eti_nomprovee.Text = "";
                            MessageBox.Show("No existe el proveedor");
                        }
                    }
                }
            }
        }
        private void Tx_nomProv_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "" && Tx_nomProv.Text.Trim() == "")
            {
                Oprove.cuenta = "";
                Oprove.codigo = "";
                Oprove.nombre = "";
                Oprove.ruc = "";
                tx_dat_provee.Text = "";
                tx_provee.Text = "";
                tx_prov.Text = "";
            }
        }
        private void xxx()
        {
            Tx_ctaDes.BackColor = ColorTranslator.FromHtml(OColores.Fondo_suave);
            Tx_ctaDes.ForeColor = tx_idOper.ForeColor;
            if (Tx_ctaDes.Text.Trim() != "")
            {
                string nt;
                DateTime fect = selecFecha1.Value.Date.AddDays(1);
                string fecOp = fect.ToString("yyyy-MM-dd"); // selecFecha1.Value.Date.ToShortDateString();
                DataRow[] row;
                if (rb_omg.Checked == true)
                {
                    // row = Program.dt_definic.Select("idtabella='DES' and descrizionerid='" + Tx_ctaDes.Text.Trim() + "'");
                    row = Program.dt_definic.Select("idtabella='DES' and descrizione='" + Tx_ctaDes.Text.Trim() + "'");
                    nt = "cassaomg";
                }
                else
                {
                    // row = Program.dt_definic.Select("idtabella='CON' and descrizionerid='" + Tx_ctaDes.Text.Trim() + "'");
                    row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + Tx_ctaDes.Text.Trim() + "'");
                    nt = "cassaconti";
                }
                if (row.Length > 0)
                {
                    eti_nomCaja.Text = row[0].ItemArray[3].ToString();  // row[0].ItemArray[2].ToString()
                    Ocajd.codigo = row[0].ItemArray[1].ToString();
                    Ocajd.largo = Tx_ctaDes.Text;      // Ocajd.nombre = Tx_ctaDes.Text
                    Ocajd.nombre = eti_nomCaja.Text.ToUpper();     // Ocajd.largo = eti_nomCaja.Text
                                                                   // vemos si la cuenta tiene saldo, si es <= 0 pone el campo en rojo
                    using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                    {
                        try
                        {
                            conn.Open();
                        }
                        catch (MySqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Application.Exit();
                            return;
                        }
                        if (conn.State == ConnectionState.Open)
                        {
                            using (MySqlCommand mic = new MySqlCommand("reps_saldoIniS", conn))  // reps_saldoIni
                            {
                                mic.CommandType = CommandType.StoredProcedure;  // retorna valor en dólares
                                mic.Parameters.AddWithValue("v_tabla", nt);
                                mic.Parameters.AddWithValue("v_a01", "LIM");
                                mic.Parameters.AddWithValue("v_a02", "");
                                mic.Parameters.AddWithValue("v_a03", "");
                                mic.Parameters.AddWithValue("v_a02f", fecOp);
                                mic.Parameters.AddWithValue("v_a04", "");
                                mic.Parameters.AddWithValue("v_a05", Ocajd.codigo);
                                using (MySqlDataReader dr = mic.ExecuteReader())
                                {
                                    if (dr.HasRows)
                                    {
                                        if (dr.Read())
                                        {
                                            if (!String.IsNullOrEmpty(dr[0].ToString()))
                                            {
                                                // Ojo, que el valor retornado esta en dólares
                                                double valor = 0;
                                                if (String.IsNullOrEmpty(Omonto.monDolar.ToString())) valor = dr.GetDouble(2);
                                                else valor = dr.GetDouble(2) - (double)Omonto.monDolar;
                                                //MessageBox.Show(valor.ToString("#0.00"), "Saldo Actual US$");
                                                if (valor <= 0)
                                                {
                                                    //Tx_ctaDes.BackColor = Color.IndianRed;    28/11/2024
                                                    Tx_ctaDes.ForeColor = System.Drawing.Color.DarkRed;
                                                }
                                            }
                                            else MessageBox.Show("No existen datos de" + Environment.NewLine +
                                                "saldo en la cuenta", "Valores nulos");
                                            tx_provee.Focus();
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("No se encuentra datos de la cuenta" + Environment.NewLine +
                                            "para calcular su saldo", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    Tx_ctaDes.Clear();
                    eti_nomCaja.Text = "";
                    MessageBox.Show("No existe el nombre de la cuenta");
                }
            }
        }
        private void yyy()
        {
            DataRow[] row;
            {
                // 21/02/2025 se agrego and numero=1
                row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + tx_ctaGiro.Text.Trim() + "' and numero=1");
            }
            if (row.Length > 0)
            {
                eti_nomCtaGiro.Text = row[0].ItemArray[3].ToString();   // nombre corto
                //Ogiro.codigo = row[0].ItemArray[1].ToString();
                Ogiro.ctades = row[0].ItemArray[3].ToString(); // Tx_ctaDes.Text;
                Ogiro.largo = row[0].ItemArray[2].ToString();   // eti_nomCtaGiro.Text
                tx_dat_giro.Text = row[0].ItemArray[1].ToString();
                Ogiro.tipodes = (rb_omg.Checked == true) ? "OMG" : "PER";
                //Ogiro.ctades = tx_dat_giro.Text;
                Ogiro.idcod = tx_dat_giro.Text;
                //Ogiro.largo = eti_nomCtaGiro.Text;
            }
            else
            {
                tx_ctaGiro.Clear();
                eti_nomCtaGiro.Text = "";
                MessageBox.Show("No existe el nombre de la cuenta");
            }
        }
        private void tx_descrip_Enter(object sender, EventArgs e)
        {
            // sino colocamos esto se va a autoseleccionar todo el texto y al dar <enter> se borrará
            // tx_descrip.SelectionStart = tx_descrip.Text.Length;  // ya no .... 24/01/2025
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
            // el codigo se traslado a changecomitted
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
                        calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), decimal.Parse(tx_tipcam.Text));
                    }
                    if (Omone.codigo == codSol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monSoles = decimal.Parse(tx_monto.Text); // Omonto.monDolar * Omonto.tipCOri;
                        Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                        calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCDol);
                    }
                    if (Omone.codigo == codEur)
                    {
                        Omonto.tipCDol = 0;
                        Omonto.tipCOri = tcDia.tcE;
                        Omonto.monEuros = decimal.Parse(tx_monto.Text);
                        Omonto.monDolar = 0;
                        Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                        calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCOri);
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
                        calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), decimal.Parse(tx_tipcam.Text));
                    }
                    if (Omone.codigo == codSol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monSoles = decimal.Parse(tx_monto.Text); // Omonto.monDolar * Omonto.tipCOri;
                        Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                        calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCDol);
                    }
                    if (Omone.codigo == codEur)
                    {
                        Omonto.tipCDol = 0;
                        Omonto.tipCOri = tcDia.tcE;
                        Omonto.monEuros = decimal.Parse(tx_monto.Text);
                        Omonto.monDolar = 0;
                        Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                        calc_monedas(cmb_mon, decimal.Parse(tx_monto.Text), Omonto.tipCOri);
                    }
                }
            }
        }
        #endregion

        #region datagridview - Grilla
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
                            consulta = "ConEgre_cassaOmg";
                        }
                        if (ntabla == "cassaconti")
                        {
                            consulta = "ConEgre_cassaConti";
                        }
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@Vdias", dAtras);
                            micon.Parameters.AddWithValue("@Vanno", 0);
                            micon.Parameters.AddWithValue("@Vidmov", "");
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_grillaE.Clear();
                                dt_grillaE.Columns.Clear();
                                da.Fill(dt_grillaE);
                                advancedDataGridView1.DataSource = dt_grillaE;
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
        }            // muestra datos de la fecha actual hasta <dAtras> días atras 
        private void armaGrilla(AdvancedDataGridView dgv_, int filasLim) // DataGridView dgv_, int filasLim
        {
            if (dgv_.Rows.Count > 0)
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
                        if (dgv_.Rows[0].Cells[i].Value != null)
                        {
                            _ = decimal.TryParse(dgv_.Rows[0].Cells[i].Value.ToString(), out decimal vd);
                            if (vd != 0) dgv_.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        }
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
        private void advancedadvancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (true)    // 24/01/2025  .... (Tx_modo.Text != "NUEVO")
            {
                string annOp = "";
                string fecOp = "";              // fecha de operacion
                decimal tipca = 0;              // tip cambio del monto origen
                string descr = "";              // descripcion de la operacion
                string idmov = "";              // id del movimiento
                if (rb_omg.Checked == true)
                {
                    // ANNO,ID_MOVIM,FECHA,DESTINO,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,idgiroconto,
                    // CTA_DESTINO,usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDDestino,IDCategoria,codimon,nombmon,TCMonOri,
                    // DET_DESTINO,DET_EGRESO,tipodesgiro,CodGiro,CTA_GIRO,CASA
                    if (Tx_modo.Text != "NUEVO") fecOp = advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);   // 24/01/2025
                    OcatEg.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDCategoria"].Value.ToString();
                    OcatEg.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["EGRESO"].Value.ToString();
                    OcatEg.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_EGRESO"].Value.ToString();
                    Omone.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = advancedDataGridView1.Rows[e.RowIndex].Cells["MONEDA"].Value.ToString();
                    Omone.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["MONTO"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoDU"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoSU"].Value.ToString());
                    tipca = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Ocajd.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDDestino"].Value.ToString();
                    Ocajd.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["DESTINO"].Value.ToString();
                    Ocajd.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_DESTINO"].Value.ToString();
                    Oprove.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["idanagrafica"].Value.ToString();
                    Oprove.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["PROVEEDOR"].Value.ToString();
                    descr = advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    if (Tx_modo.Text != "NUEVO") idmov = advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString(); // 24/01/2025
                    Ogiro.ctades = advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_GIRO"].Value.ToString();
                    Ogiro.tipodes = advancedDataGridView1.Rows[e.RowIndex].Cells["tipodesgiro"].Value.ToString();
                    Ogiro.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["CodGiro"].Value.ToString();
                    Ogiro.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_DESTINO"].Value.ToString();
                    Ogiro.idcod = advancedDataGridView1.Rows[e.RowIndex].Cells["IDGiroConto"].Value.ToString();
                    annOp = advancedDataGridView1.Rows[e.RowIndex].Cells["ANNO"].Value.ToString();
                }
                else
                {
                    // ANNO,ID_MOVIM,FECHA,CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,
                    // GIRO_CTA,IDGiroConto,CTA_DESTINO,usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDConto,
                    // IDCategoria,codimon,nombmon,TCMonOri,DET_CUENTA,DET_EGRESO,CodGiro,CASA,CTA_GIRO
                    if (Tx_modo.Text != "NUEVO") fecOp = advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);
                    OcatEg.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDCategoria"].Value.ToString();
                    OcatEg.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["EGRESO"].Value.ToString();
                    OcatEg.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_EGRESO"].Value.ToString();
                    Omone.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = advancedDataGridView1.Rows[e.RowIndex].Cells["MONEDA"].Value.ToString();
                    Omone.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["MONTO"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoDU"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoSU"].Value.ToString());
                    tipca = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Ocajd.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDConto"].Value.ToString();
                    Ocajd.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["CUENTA"].Value.ToString();
                    Ocajd.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_CUENTA"].Value.ToString();
                    Oprove.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["idanagrafica"].Value.ToString();
                    Oprove.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["PROVEEDOR"].Value.ToString();
                    descr = advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    if (Tx_modo.Text != "NUEVO") idmov = advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString(); // 24/01/2025
                    Ogiro.ctades = advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_GIRO"].Value.ToString();
                    Ogiro.tipodes = advancedDataGridView1.Rows[e.RowIndex].Cells["GIRO_CTA"].Value.ToString();
                    Ogiro.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["CodGiro"].Value.ToString();
                    Ogiro.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_DESTINO"].Value.ToString();
                    Ogiro.idcod = advancedDataGridView1.Rows[e.RowIndex].Cells["IDGiroConto"].Value.ToString();
                    annOp = advancedDataGridView1.Rows[e.RowIndex].Cells["ANNO"].Value.ToString();
                }
                Oegreso.creaEgreso(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, tipca,
                        Ocajd, Oprove, descr, idmov, Ogiro, annOp);
                jalaoc();
            }
        }
        private void insFilaEnDataG(string _casa, string _corre)
        {
            DataRow fila = dt_grillaE.NewRow();
            string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
            advancedDataGridView1.Rows[0].DefaultCellStyle.BackColor = System.Drawing.SystemColors.Window;  // 22/04/2025
            if (rb_omg.Checked == true)
            {
                // CASA,AÑO,ID_MOVIM,FECHA,DESTINO,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,idgiroconto,CTA_DESTINO,
                // usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_EGRESO,
                // tipodesgiro,CodGiro,CTA_GIRO
                fila["CASA"] = _casa;
                fila["ANNO"] = tx_anno.Text;
                fila["ID_MOVIM"] = _corre;
                fila["FECHA"] = fecOp;
                fila["DESTINO"] = Ocajd.nombre;     // nombre cuenta destino
                fila["EGRESO"] = OcatEg.nombre.ToUpper();     // nombre categoria egreso
                fila["MONEDA"] = Omone.siglas;      // siglas moneda origen
                fila["MONTO"] = Omonto.monOrige;    // valor origen
                fila["DESCRIPCION"] = tx_descrip.Text;
                fila["TIP_CAMBIO"] = Omonto.tipCOri; // decimal.Parse(tx_tipcam.Text);
                fila["PROVEEDOR"] = Oprove.nombre;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["idgiroconto"] = Ogiro.idcod;
                fila["CTA_DESTINO"] = tx_ctaGiro.Text; // Ogiro.largo;
                fila["usuario"] = Program.vg_user;
                //fila["dia"] = "";
                fila["ImportoDU"] = Omonto.monDolar;
                fila["ImportoSU"] = Omonto.monSoles;
                fila["idanagrafica"] = Oprove.codigo;
                fila["IDDestino"] = Ocajd.codigo;
                fila["IDCategoria"] = OcatEg.codigo;
                fila["codimon"] = Omone.codigo;
                fila["nombmon"] = Omone.nombre;
                fila["TCMonOri"] = tx_tipcam.Text;
                fila["DET_DESTINO"] = Ocajd.largo;
                fila["DET_EGRESO"] = OcatEg.largo;
                fila["tipodesgiro"] = Ogiro.tipodes;
                fila["CodGiro"] = Ogiro.codigo;
                fila["CTA_GIRO"] = Ogiro.ctades;
            }
            if (rb_pers.Checked == true)
            {
                // CASA,AÑO,ID_MOVIM,FECHA,CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,IDGiroConto,CTA_DESTINO,usuario,
                //dia,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,CTA_GIRO,codimon,nombmon,TCMonOri,DET_CUENTA,DET_EGRESO,tipodesgiro,CodGiro
                fila["CASA"] = _casa;
                fila["ANNO"] = tx_anno.Text;
                fila["ID_MOVIM"] = _corre;
                fila["FECHA"] = fecOp;
                fila["CUENTA"] = Ocajd.nombre;
                fila["EGRESO"] = OcatEg.nombre.ToUpper();
                fila["MONEDA"] = Omone.siglas;
                fila["MONTO"] = Omonto.monOrige;
                fila["DESCRIPCION"] = tx_descrip.Text;
                fila["TIP_CAMBIO"] = Omonto.tipCOri; //  decimal.Parse(tx_tipcam.Text);
                fila["PROVEEDOR"] = Oprove.nombre;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["IDGiroConto"] = Ogiro.idcod;
                fila["CTA_DESTINO"] = tx_ctaGiro.Text; // Ogiro.largo;
                fila["usuario"] = Program.vg_user;
                //fila["dia"] = "";
                fila["ImportoDU"] = Omonto.monDolar;
                fila["ImportoSU"] = Omonto.monSoles;
                fila["idanagrafica"] = Oprove.codigo;
                fila["IDConto"] = Ocajd.codigo;
                fila["IDCategoria"] = OcatEg.codigo;
                fila["CTA_GIRO"] = Ogiro.ctades;
                fila["codimon"] = Omone.codigo;
                fila["nombmon"] = Omone.nombre;
                fila["TCMonOri"] = Omonto.tipCOri;
                fila["DET_CUENTA"] = Ocajd.largo;
                fila["DET_EGRESO"] = OcatEg.largo;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["CodGiro"] = Ogiro.codigo;
            }
            dt_grillaE.Rows.InsertAt(fila, 0);
            advancedDataGridView1.CurrentCell = advancedDataGridView1.Rows[0].Cells[0];
            advancedDataGridView1.CurrentRow.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(col1rafila);   // 22/04/2025
        }                // INSERTA en la grilla el registro nuevo despues de grabar en la B.D.
        public void actFilaEnDataG(DataTable dt, string _casa, string _corre)
        {
            string fecOp = Tx_fecha.Text;  // selecFecha1.Value.Date.ToShortDateString();
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = dt.Rows[i];
                if (dr["ANNO"].ToString() == tx_anno.Text && dr["ID_MOVIM"].ToString() == CDerecha(_corre, 6))  //    dr["ID_MOVIM"].ToString() == (_corre.Substring(0, 4) + CDerecha(_corre, 6))
                {
                    if (rb_omg.Checked == true)
                    {
                        // CASA,AÑO,ID_MOVIM,FECHA,DESTINO,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,idgiroconto,CTA_DESTINO,
                        // usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_EGRESO,
                        //tipodesgiro,CodGiro,CTA_GIRO
                        dr["CASA"] = _casa;
                        dr["ANNO"] = tx_anno.Text;
                        dr["ID_MOVIM"] = _corre;
                        dr["FECHA"] = fecOp;
                        dr["DESTINO"] = Ocajd.nombre;     // nombre cuenta destino
                        dr["EGRESO"] = OcatEg.nombre.ToUpper();     // nombre categoria egreso
                        dr["MONEDA"] = Omone.siglas;      // siglas moneda origen
                        dr["MONTO"] = Omonto.monOrige;    // valor origen
                        dr["DESCRIPCION"] = tx_descrip.Text;
                        dr["TIP_CAMBIO"] = Omonto.tipCOri; // decimal.Parse(tx_tipcam.Text);
                        dr["PROVEEDOR"] = Oprove.nombre;
                        dr["GIRO_CTA"] = Ogiro.tipodes;
                        dr["idgiroconto"] = Ogiro.idcod;
                        dr["CTA_DESTINO"] = Ogiro.largo;
                        dr["usuario"] = Program.vg_user;
                        //dr["dia"] = "";
                        dr["ImportoDU"] = Omonto.monDolar;
                        dr["ImportoSU"] = Omonto.monSoles;
                        dr["idanagrafica"] = Oprove.codigo;
                        dr["IDDestino"] = Ocajd.codigo;
                        dr["IDCategoria"] = OcatEg.codigo;
                        dr["codimon"] = Omone.codigo;
                        dr["nombmon"] = Omone.nombre;
                        dr["TCMonOri"] = tx_tipcam.Text;
                        dr["DET_DESTINO"] = Ocajd.largo;
                        dr["DET_EGRESO"] = OcatEg.largo;
                        dr["tipodesgiro"] = Ogiro.tipodes;
                        dr["CodGiro"] = Ogiro.codigo;
                        dr["CTA_GIRO"] = Ogiro.ctades;
                    }
                    if (rb_pers.Checked == true)
                    {
                        // CASA,AÑO,ID_MOVIM,FECHA,CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,IDGiroConto,CTA_DESTINO,usuario,
                        //dia,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,CTA_GIRO,codimon,nombmon,TCMonOri,DET_CUENTA,DET_EGRESO,tipodesgiro,CodGiro
                        dr["CASA"] = _casa;
                        dr["ANNO"] = tx_anno.Text;
                        dr["ID_MOVIM"] = _corre;
                        dr["FECHA"] = fecOp;
                        dr["CUENTA"] = Ocajd.nombre;
                        dr["EGRESO"] = OcatEg.nombre.ToUpper();
                        dr["MONEDA"] = Omone.siglas;
                        dr["MONTO"] = Omonto.monOrige;
                        dr["DESCRIPCION"] = tx_descrip.Text;
                        dr["TIP_CAMBIO"] = Omonto.tipCOri; //  decimal.Parse(tx_tipcam.Text);
                        dr["PROVEEDOR"] = Oprove.nombre;
                        dr["GIRO_CTA"] = Ogiro.tipodes;
                        dr["IDGiroConto"] = Ogiro.idcod;
                        dr["CTA_DESTINO"] = Ogiro.largo;
                        dr["usuario"] = Program.vg_user;
                        //dr["dia"] = "";
                        dr["ImportoDU"] = Omonto.monDolar;
                        dr["ImportoSU"] = Omonto.monSoles;
                        dr["idanagrafica"] = Oprove.codigo;
                        dr["IDConto"] = Ocajd.codigo;
                        dr["IDCategoria"] = OcatEg.codigo;
                        dr["codimon"] = Omone.codigo;
                        dr["nombmon"] = Omone.nombre;
                        dr["TCMonOri"] = Omonto.tipCOri;
                        dr["DET_CUENTA"] = Ocajd.largo;
                        dr["DET_EGRESO"] = OcatEg.largo;
                        dr["GIRO_CTA"] = Ogiro.tipodes;
                        dr["CodGiro"] = Ogiro.codigo;
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
            Bt_graba.BackColor = System.Drawing.Color.PeachPuff; // ColorTranslator.FromHtml("#fabdba");
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
                if (Tx_catEgre.Text == "")
                {
                    errorProvider1.SetError(Tx_catEgre, "Debe ingresar un tipo");
                    Tx_catEgre.Focus();
                    return;
                }
                errorProvider1.SetError(Tx_catEgre, "");
                if (Tx_ctaDes.Text == "")
                {
                    errorProvider1.SetError(Tx_ctaDes, "Debe seleccionar la cuenta");
                    Tx_ctaDes.Focus();
                    return;
                }
                errorProvider1.SetError(Tx_ctaDes, "");
                if (tx_tipcam.Text.Trim() == "0" || tx_tipcam.Text.Trim() == "")
                {
                    errorProvider1.SetIconAlignment(tx_tipcam, ErrorIconAlignment.TopLeft);
                    errorProvider1.SetError(tx_tipcam, "Debe ingresar el tipo de cambio");
                    tx_tipcam.Focus();
                    return;
                }
                errorProvider1.SetError(tx_tipcam, "");
                if (cmb_mon.Text == "")
                {
                    errorProvider1.SetError(cmb_mon, "Debe seleccionar la moneda");
                    cmb_mon.Focus();
                    return;
                }
                errorProvider1.SetError(cmb_mon, "");
                if (tx_monto.Text == "")
                {
                    errorProvider1.SetError(tx_monto, "Debe ingresar un valor");
                    tx_monto.Focus();
                    return;
                }
                errorProvider1.SetError(tx_monto, "");
                if (chk_giroC.CheckState == CheckState.Checked)
                {
                    if (tx_ctaGiro.Text.Trim() == "")
                    {
                        MessageBox.Show("Debe ingresar la cuenta destino del giro", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        tx_ctaGiro.Focus();
                        return;
                    }
                    if (String.IsNullOrEmpty(Ogiro.idcod) || String.IsNullOrEmpty(Ogiro.ctades) || String.IsNullOrEmpty(Ogiro.largo))
                    {
                        MessageBox.Show("Complete la cuenta del Giro", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        tx_ctaGiro.Select();
                        return;
                    }
                    if (eti_nomCtaGiro.Text.Trim() != Ogiro.ctades)  // 27/01/2025
                    {
                        MessageBox.Show("Cuenta de Giro no completa", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        tx_ctaGiro.Select();
                        return;
                    }
                }
                // ******************* hoy 13/12/2024 
                if (String.IsNullOrEmpty(OcatEg.codigo) || String.IsNullOrEmpty(OcatEg.nombre) || String.IsNullOrEmpty(OcatEg.largo))
                {
                    MessageBox.Show("Complete la categoría de Egreso", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_catEgre.Focus();
                    return;
                }
                if (Tx_catEgre.Text.Trim().ToUpper() != OcatEg.largo.ToUpper())
                {
                    MessageBox.Show("Categoría de Egreso no completa", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_catEgre.Focus();
                    return;
                }
                if (String.IsNullOrEmpty(Ocajd.codigo) || String.IsNullOrEmpty(Ocajd.nombre) || String.IsNullOrEmpty(Ocajd.largo))
                {
                    MessageBox.Show("Complete la cuenta de destino", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_ctaDes.Select();
                    return;
                }
                if (Tx_ctaDes.Text.Trim() != Ocajd.largo)
                {
                    MessageBox.Show("Cuenta destino no completa", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_ctaDes.Select();
                    return;
                }
                if (Omone.codigo == "" || Omone.nombre == "")
                {
                    MessageBox.Show("Seleccione el tipo de moneda", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    cmb_mon.Focus();
                    return;
                }
                if (Omonto.codMOrige == "" || Omonto.monSoles <= 0)
                {
                    MessageBox.Show("Ingrese el importe", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    tx_monto.Focus();
                    return;
                }
                // **********************************
                graba_nuevo();
                Tx_catEgre.Focus();
            }
            if (Tx_modo.Text == "EDICION")
            {
                if (tx_idOper.Text == "")
                {
                    MessageBox.Show("No hay registro que Editar!", "Identificador en blanco", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }
                var aaa = MessageBox.Show("Confirma que desea EDITAR el Egreso?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    graba_edicion(dt_grillaE);
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
                var aaa = MessageBox.Show("Confirma que desea BORRAR el Egreso?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    string tabla = "";
                    if (rb_omg.Checked == true) tabla = "cassaomg";
                    else tabla = "cassaconti";
                    dt_grillaE.TableName = "dt_grillaE";
                    graba_borrar(tabla, selecFecha1.Value.Year.ToString(), "000000000" + CDerecha(tx_idOper.Text, 6), dt_grillaE);
                    limpiaObj();
                    limpiaTE();
                }
            }
        }
        private void graba_nuevo()
        {
            var aaa = MessageBox.Show("Confirma que desea crear el Egreso?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (aaa == DialogResult.Yes)
            {
                string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
                Egresos Oegresos = new Egresos();
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
                        if (true)   // Tx_modo.Text == "NUEVO"
                        {
                            string corre = correlativo(conn, ((rb_omg.Checked == true) ? "MCA" : "MCO"), int.Parse(tx_anno.Text));
                            string corrA = corre;
                            if (corre != "error" && corre != "")
                            {
                                try
                                {
                                    Oegresos.creaEgreso(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                                        Ocajd, Oprove, tx_descrip.Text, corre, Ogiro, tx_anno.Text);
                                    Oegresos.grabaEgreso(conn);
                                    // si esta marcado el giro, hacemos el movimiento inverso
                                    if (chk_giroC.CheckState == CheckState.Checked)
                                    {
                                        oper_giro(conn, Ogiro, (rb_omg.Checked == true) ? "cassaomg" : "cassaconti", OcatEg, fecOp, Omone, Omonto, decimal.Parse(tx_tipcam.Text), tx_descrip.Text, Ocajd.codigo, tx_anno.Text);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Error en grabar Egreso");
                                    return;
                                }
                                insFilaEnDataG("LIM", CDerecha("00000" + corrA, 6));       // inserta el registro nuevo en la grilla
                                if (chk_datSimil.CheckState == CheckState.Checked)
                                {
                                    datsimil();
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
        public void graba_edicion(DataTable dgv)
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
                        string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
                        string corre = tx_anno.Text + tx_idOper.Text;
                        
                        Oegreso.creaEgreso(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                                        Ocajd, Oprove, tx_descrip.Text, corre, Ogiro, tx_anno.Text);
                        Oegreso.EditaEgreso(conn, tx_anno.Text, ("000000000" + CDerecha(tx_idOper.Text, 6)));
                        if (chk_giroC.CheckState == CheckState.Checked)
                        {
                            catIngresos OcatIn = new catIngresos();
                            OcatIn.codigo = OcatEg.codigo;
                            OcatIn.nombre = OcatEg.nombre;
                            OcatIn.largo = OcatEg.largo;
                            cajDestino _desgiro = new cajDestino();
                            _desgiro.codigo = tx_dat_giro.Text;
                            _desgiro.nombre = tx_ctaGiro.Text;
                            _desgiro.largo = eti_nomCtaGiro.Text;
                            // jalamos el idmov a partir del codigo del giro
                            Ogiro.idcod = Ocajd.codigo;
                            corre = jalaIDMov(conn , Ogiro.codigo, ((rb_omg.Checked == true) ? "cassaomg" : "cassaconti"), Ogiro.idcod);
                            Oingresos.creaIngreso(pan_p.Tag.ToString(), fecOp, OcatIn, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                            _desgiro, tx_descrip.Text, corre, Ogiro, tx_anno.Text);
                            Oingresos.EditaIngreso(conn, tx_anno.Text, corre);   // "000000000" + CDerecha(tx_idOper.Text, 6)
                        }
                        actFilaEnDataG(dt_grillaE, "LIM", tx_idOper.Text);
                        limpiaObj();
                        limpiaTE();
                    }
                }
            }
        }
        public void graba_borrar(string tabla, string year, string idmov, DataTable dgv)    // OJO, este idmov debe incluir todos los ceros a la izquierda
        {
            if (true)
            {
                // buscamos si tiene giroconto
                string _giro_ = "no";  // por defecto asumimos que no tiene giro
                if (tabla != "camion" && dgv.TableName == "dt_grillaE")  // giroconton solo para EGRESOS 25/09/2024
                {
                    for (int i = dgv.Rows.Count - 1; i >= 0; i--)
                    {
                        DataRow dr = dgv.Rows[i];
                        if (dr["ANNO"].ToString() == year && dr["ID_MOVIM"].ToString() == CDerecha(idmov, 6))  // (year + CDerecha(idmov, 6)))
                        {
                            if (dr["CodGiro"].ToString().Trim() != "")
                            {
                                var aaa = MessageBox.Show("El registro tiene GIROCONTO, desea" + Environment.NewLine +
                                    "BORRAR la cuenta también?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (aaa == DialogResult.Yes)
                                {
                                    _giro_ = dr["CodGiro"].ToString();
                                }
                            }
                        }
                    }
                }
                // borra en la tabla
                // borra en la grilla
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
                        string consulta = ""; 
                        if (_giro_ == "no") consulta = "delete from " + tabla + " where anno=@year and idmovimento=@corre";
                        if (_giro_ != "no")
                        {
                            consulta = "delete from " + tabla + " where CodGiro=@codG";
                            // y que pasaría si el giroconto fuera entre las dos tblas cassaomg y cassaconti ????? 07/09/2024
                        }
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            if (_giro_ != "no")
                            {
                                micon.Parameters.AddWithValue("@codG", _giro_);
                            }
                            else
                            {
                                micon.Parameters.AddWithValue("@year", year);
                                micon.Parameters.AddWithValue("@corre", idmov);
                            }
                            micon.ExecuteNonQuery();
                        }
                        for (int i = dgv.Rows.Count - 1; i >= 0; i--)
                        {
                            DataRow dr = dgv.Rows[i];   //(year + CDerecha(idmov, 6)))
                            if (dr["ID_MOVIM"].ToString() == CDerecha(idmov, 6) && dr["ANNO"].ToString() == year)
                                dr.Delete();
                        }
                        dgv.AcceptChanges();
                    }
                }
            }
        }
        private void bt_Pnuevo_Click(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" && bt_Pnuevo.Enabled == true)
            {
                string para1 = "provee";
                string para2 = "";
                string para3 = "";
                string para4 = "";    // todos | activos
                ayuda1 ayu1 = new ayuda1(para1, para2, para3, para4);
                var result = ayu1.ShowDialog();
                    if (result == DialogResult.Cancel)
                {
                    if (!string.IsNullOrEmpty(ayu1.ReturnValue1))
                    {
                        tx_dat_provee.Text = ayu1.ReturnValue0;
                        Tx_nomProv.Text = ayu1.ReturnValue1;
                        tx_prov.Text = ayu1.ReturnValue1;
                        eti_nomprovee.Text = ayu1.ReturnValue1; // 31/08/2024 ya no usamos
                        Oprove.codigo = ayu1.ReturnValueA[0];
                        Oprove.nombre = ayu1.ReturnValueA[1];
                        Oprove.ruc = ayu1.ReturnValueA[2];
                        Oprove.cuenta = ayu1.ReturnValueA[3];
                        
                        DataRow dr = dtpro.NewRow();
                        dr[0] = Oprove.codigo;
                        dr[1] = Oprove.nombre.Replace("\r\n", string.Empty);
                        dtpro.Rows.Add(dr);
                        dtpro.AcceptChanges();
                        lista_prov.Add(Oprove.nombre);
                        Tx_nomProv.Values = lista_prov.ToArray();
                        SendKeys.Send("{Tab}");
                    }
                }
            }
        }
        #endregion

        public montos calc_monedas(ComboBox combo, decimal valOri, decimal tipCam)
        {
            if (valOri <= 0) return Omonto;
            if (tipCam <= 0) return Omonto;
            if (combo.SelectedValue == null) return Omonto;
            Omonto.codMOrige = combo.SelectedValue.ToString();              // codigo de la moneda
            Omonto.monOrige = valOri;
            if (combo.SelectedValue.ToString() == codSol) // Soles
            {
                Omonto.monSoles = valOri;
                Omonto.tipCDol = tipCam;
                Omonto.monDolar = Math.Round((valOri / tipCam), 2);
                Omonto.tipCOri = tipCam;
            }
            if (combo.SelectedValue.ToString() == codDol) // Dolares
            {
                Omonto.tipCDol = tipCam;
                Omonto.monDolar = valOri;
                Omonto.monSoles = Math.Round((valOri * tipCam), 2);
                Omonto.tipCOri = tipCam;
            }
            if (combo.SelectedValue.ToString() == codEur) // Euros
            {
                Omonto.tipCDol = 0;
                Omonto.monEuros = valOri;
                Omonto.tipCOri = tipCam;
                Omonto.monSoles = Math.Round((valOri * tipCam), 2);
            }
            return Omonto;
        }
        public string correlativo(MySqlConnection conn, string idcont, int year)
        {
            string retorna = "";
            int contador = 0;
            string consulta = "select numero from contatori where idbanco='LIM' and anno=@year and idcontatore=@idcont";
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                micon.Parameters.AddWithValue("@year", year);
                micon.Parameters.AddWithValue("@idcont", idcont);
                using (MySqlDataReader dr = micon.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            contador = dr.GetInt32(0) + 1;
                            retorna = CDerecha("00000000000000" + contador.ToString(), 15);
                        }
                    }
                    else
                    {
                        retorna = "error";
                    }
                }
            }
            if (retorna != "error" && retorna != "")
            {
                using (MySqlCommand micon = new MySqlCommand("update contatori set numero=@contador where idbanco='LIM' and anno=@year and idcontatore=@idcont", conn))
                {
                    micon.Parameters.AddWithValue("@year", year);
                    micon.Parameters.AddWithValue("@idcont", idcont);
                    micon.Parameters.AddWithValue("@contador", contador);
                    micon.ExecuteNonQuery();
                }
            }
            return retorna;
        }
        public string CDerecha(string sValue, int iMaxLength)
        {
            if (string.IsNullOrEmpty(sValue))
            {
                sValue = string.Empty;
            }
            else if (sValue.Length > iMaxLength)
            {
                sValue = sValue.Substring(sValue.Length - iMaxLength, iMaxLength);
            }
            return sValue;
        }                  // devuelve los ultimos n caractares desde la derecha
        public string[] jala_ultimo(DataTable dt_G, string tipo, string tcuenta, string fecha)
        {
            // los datos deben jalarse de la grilla actual porque ahi estan los datos! 
            string[] retorna = new string[21];
            retorna[0] = "";  // tipo de cambio
            retorna[1] = "";  // categoria id
            retorna[2] = "";  // categoria corto  
            retorna[3] = "";  // categoria largo
            retorna[4] = "";  // moneda codigo
            retorna[5] = "";  // moneda nombre
            retorna[6] = "";  // cuenta id
            retorna[7] = "";  // cuenta corto
            retorna[8] = "";  // cuenta largo
            retorna[9] = "";  // proveedor id
            retorna[10] = ""; // proveedor nombre
            retorna[11] = ""; // descripción
            retorna[12] = ""; // ctaGiro id
            retorna[13] = ""; // ctaGiro corto
            retorna[14] = ""; // ctaGiro largo
            retorna[15] = "";   // moneda siglas
            retorna[16] = "";   // monto original
            retorna[17] = "";   // monto en dolares
            retorna[18] = "";   // monto soles
            retorna[19] = "";   // fecha de la operación
            retorna[20] = "";   // año de la operación

            DataRow[] row = null;
            try
            {
                row = dt_G.Select("FECHA='" + fecha + "'", "ID_MOVIM DESC");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error en datos", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return retorna;
            }
            if (row != null && row.Length > 0)
            {
                if (tcuenta == "OMG")
                {
                    if (tipo == "EGRESO")
                    {
                        // ANNO,ID_MOVIM,FECHA,DESTINO,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,idgiroconto,
                        // CTA_DESTINO,usuario,dia,ImportoDU,ImportoSU,idanagrafica,IDDestino,IDCategoria,codimon,nombmon,TCMonOri,
                        // DET_DESTINO,DET_EGRESO,tipodesgiro,CodGiro,CTA_GIRO,CASA
                        retorna[0] = row[0]["TIP_CAMBIO"].ToString();   // tipo de cambio
                        retorna[1] = row[0]["IDCategoria"].ToString();  // categoria id
                        retorna[2] = row[0]["EGRESO"].ToString();       // categoria corto  
                        retorna[3] = row[0]["DET_EGRESO"].ToString();   // categoria largo
                        retorna[4] = row[0]["codimon"].ToString();      // moneda codigo
                        retorna[5] = row[0]["nombmon"].ToString();      // moneda nombre
                        retorna[6] = row[0]["IDDestino"].ToString();    // cuenta id
                        retorna[7] = row[0]["DESTINO"].ToString();      // cuenta corto
                        retorna[8] = row[0]["DET_DESTINO"].ToString();  // cuenta largo
                        retorna[9] = row[0]["idanagrafica"].ToString(); // proveedor id
                        retorna[10] = row[0]["PROVEEDOR"].ToString();   // proveedor nombre
                        retorna[11] = row[0]["DESCRIPCION"].ToString(); // descripción
                        retorna[12] = row[0]["IDGiroConto"].ToString(); // ctaGiro id      | no debemos
                        retorna[13] = row[0]["CTA_GIRO"].ToString();    // ctaGiro corto   | jalar el (CodGiro)
                        retorna[14] = row[0]["CTA_DESTINO"].ToString(); // ctaGiro largo   | codigo del giro
                        retorna[15] = row[0]["MONEDA"].ToString();      // moneda siglas 
                        retorna[16] = row[0]["MONTO"].ToString();       // monto original
                        retorna[17] = row[0]["ImportoDU"].ToString();   // monto en dolares
                        retorna[18] = row[0]["ImportoSU"].ToString();   // monto soles
                        retorna[19] = row[0]["FECHA"].ToString();       // fecha de la oper
                        retorna[20] = row[0]["ANNO"].ToString();        // año de la oper
                    }
                    if (tipo == "INGRESO")
                    {
                        // ANNO,ID_MOVIM,FECHA,DESTINO,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,a.usuario,a.dia,ImportoDE,ImportoSE,
                        // IDDestino,IDCategoria,codimon,nombmon,TCMonOri,DET_DESTINO,DET_INGRESO,a.tipodesgiro,a.CodGiro,CTA_GIRO,
                        // CASA,GIRO_CTA,a.idgiroconto,CTA_DESTINO
                        retorna[0] = row[0]["TIP_CAMBIO"].ToString();   // tipo de cambio
                        retorna[1] = row[0]["IDCategoria"].ToString();  // categoria id
                        retorna[2] = row[0]["INGRESO"].ToString();      // categoria corto  
                        retorna[3] = row[0]["DET_INGRESO"].ToString();  // categoria largo
                        retorna[4] = row[0]["codimon"].ToString();      // moneda codigo
                        retorna[5] = row[0]["nombmon"].ToString();      // moneda nombre
                        retorna[6] = row[0]["IDDestino"].ToString();    // cuenta id
                        retorna[7] = row[0]["DESTINO"].ToString();      // cuenta corto
                        retorna[8] = row[0]["DET_DESTINO"].ToString();  // cuenta nombre largo
                        retorna[9] = "";                                // proveedor id
                        retorna[10] = "";                               // proveedor nombre
                        retorna[11] = row[0]["DESCRIPCION"].ToString(); // descripción
                        retorna[12] = "";   // row[0]["idgiroconto"].ToString(); // ctaGiro id
                        retorna[13] = "";   // row[0]["CTA_GIRO"].ToString();    // ctaGiro corto
                        retorna[14] = "";   // row[0]["CTA_DESTINO"].ToString(); // ctaGiro largo
                        retorna[15] = row[0]["MONEDA"].ToString();      // moneda siglas 
                        retorna[16] = row[0]["MONTO"].ToString();       // monto original
                        retorna[17] = row[0]["ImportoDE"].ToString();   // monto en dolares
                        retorna[18] = row[0]["ImportoSE"].ToString();   // monto soles
                        retorna[19] = row[0]["FECHA"].ToString();       // fecha de la oper
                        retorna[20] = row[0]["ANNO"].ToString();        // año de la oper
                    }
                }
                if (tcuenta == "PER")
                {
                    if (tipo == "EGRESO")
                    {
                        // CASA,AÑO,ID_MOVIM,FECHA,CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,GIRO_CTA,IDGiroConto,CTA_DESTINO,usuario,
                        // dia,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,CTA_GIRO,codimon,nombmon,TCMonOri,DET_CUENTA,DET_EGRESO,tipodesgiro,CodGiro
                        retorna[0] = row[0]["TIP_CAMBIO"].ToString();   // tipo de cambio
                        retorna[1] = row[0]["IDCategoria"].ToString();  // categoria id
                        retorna[2] = row[0]["EGRESO"].ToString();       // categoria corto  
                        retorna[3] = row[0]["DET_EGRESO"].ToString();   // categoria largo
                        retorna[4] = row[0]["codimon"].ToString();      // moneda codigo
                        retorna[5] = row[0]["nombmon"].ToString();      // moneda nombre
                        retorna[6] = row[0]["IDConto"].ToString();      // cuenta id
                        retorna[7] = row[0]["CUENTA"].ToString();       // cuenta corto
                        retorna[8] = row[0]["DET_CUENTA"].ToString();   // cuenta largo
                        retorna[9] = row[0]["idanagrafica"].ToString(); // proveedor id
                        retorna[10] = row[0]["PROVEEDOR"].ToString();   // proveedor nombre
                        retorna[11] = row[0]["DESCRIPCION"].ToString(); // descripción
                        retorna[12] = row[0]["IDGiroConto"].ToString(); // ctaGiro id      | no debemos
                        retorna[13] = row[0]["CTA_GIRO"].ToString();    // ctaGiro corto   | jalar el (CodGiro)
                        retorna[14] = row[0]["CTA_DESTINO"].ToString(); // ctaGiro largo   | codigo del giro
                        retorna[15] = row[0]["MONEDA"].ToString();      // moneda siglas 
                        retorna[16] = row[0]["MONTO"].ToString();       // monto original
                        retorna[17] = row[0]["ImportoDU"].ToString();   // monto en dolares
                        retorna[18] = row[0]["ImportoSU"].ToString();   // monto soles
                        retorna[19] = row[0]["FECHA"].ToString();       // fecha de la oper
                        retorna[20] = row[0]["ANNO"].ToString();        // año de la oper
                    }
                    if (tipo == "INGRESO")
                    {
                        // CASA,AÑO,ID_MOVIM,FECHA,CUENTA,INGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,GIRO_CTA,
                        // IDGiroConto,CTA_DESTINO,usuario,dia,ImportoDE,ImportoSE,IDConto,IDCategoria,codimon,
                        // nombmon,TCMonOri,DET_CUENTA,DET_INGRESO,tipodesgiro,CodGiro.CTA_GIRO
                        retorna[0] = row[0]["TIP_CAMBIO"].ToString();   // tipo de cambio
                        retorna[1] = row[0]["IDCategoria"].ToString();  // categoria id
                        retorna[2] = row[0]["INGRESO"].ToString();      // categoria corto  
                        retorna[3] = row[0]["DET_INGRESO"].ToString();  // categoria largo
                        retorna[4] = row[0]["codimon"].ToString();       // moneda codigo
                        retorna[5] = row[0]["nombmon"].ToString();      // moneda nombre
                        retorna[6] = row[0]["IDConto"].ToString();      // cuenta id
                        retorna[7] = row[0]["CUENTA"].ToString();       // cuenta corto
                        retorna[8] = row[0]["DET_CUENTA"].ToString();   // cuenta largo
                        retorna[9] = "";  // proveedor id
                        retorna[10] = ""; // proveedor nombre
                        retorna[11] = row[0]["DESCRIPCION"].ToString(); // descripción
                        retorna[12] = row[0]["IDGiroConto"].ToString(); // ctaGiro id
                        retorna[13] = row[0]["CTA_GIRO"].ToString(); // ctaGiro corto
                        retorna[14] = row[0]["CTA_DESTINO"].ToString(); // ctaGiro largo
                        retorna[15] = row[0]["MONEDA"].ToString();      // moneda siglas 
                        retorna[16] = row[0]["MONTO"].ToString();       // monto original
                        retorna[17] = row[0]["ImportoDE"].ToString();   // monto en dolares
                        retorna[18] = row[0]["ImportoSE"].ToString();   // monto soles
                        retorna[19] = row[0]["FECHA"].ToString();       // fecha de la oper
                        retorna[20] = row[0]["ANNO"].ToString();        // año de la oper
                    }
                }
            }
            return retorna;
        }   // jala el ultimo registro OMG/Personal, Egreso/Ingreso, Fecha
        public string jalaIDMov(MySqlConnection conn, string codGiro, string tabla, string ctaG)
        {
            string retorna = "";
            if (conn.State == ConnectionState.Open)
            {
                string aaa = "select * from " + tabla + " where CodGiro=@cg and IDGiroConto=@idg";
                using (MySqlCommand mcom = new MySqlCommand(aaa, conn))
                {
                    mcom.Parameters.AddWithValue("@cg", codGiro);       // codigo del Giro
                    mcom.Parameters.AddWithValue("@idg", ctaG);         // cta destino del Giro (idgiroconto)
                    using (MySqlDataReader dr = mcom.ExecuteReader())
                    {
                        if (dr.HasRows)
                        {
                            if (dr.Read())
                            {
                                retorna = dr.GetString(2);
                            }
                        }
                    }
                }
            }

            return retorna;
        }
        public void oper_giro(MySqlConnection conn, giroConto giro, string tipcta, catEgresos catEg, string fecOp, 
            monedas mone, montos monto, decimal tcamb, string descrip, string ccdes, string aop)  
        {
            // fecOp = fecha de la operacion formato dd/MM/aaaa
            // tipcta = "cassaomg" || "cassaconti"
            string _codGiro = "";
            string _lastIn = "";
            // jalamos el id para crear el CodGiro para grabarlo aqui
            using (MySqlCommand mic = new MySqlCommand("select CAST(last_insert_id() as int)", conn))
            {
                using (MySqlDataReader dr = mic.ExecuteReader())
                {
                    if (dr.Read())
                    {
                        _lastIn = dr.GetUInt64(0).ToString();
                        _codGiro = giro.tipodes + _lastIn;
                        giro.codigo = _codGiro;
                    }
                }
            }
            string actua = "update " + tipcta + " set CodGiro=@_codi " +
                "where " + ((tipcta == "cassaomg") ? "idcassaomg" : "idcassaconti") + " = @_id";
            using (MySqlCommand mic = new MySqlCommand(actua, conn))
            {
                mic.Parameters.AddWithValue("@_codi", _codGiro);
                mic.Parameters.AddWithValue("@_id", _lastIn);
                mic.ExecuteNonQuery();
            }

            catIngresos OcatIn = new catIngresos();
            OcatIn.codigo = catEg.codigo;
            OcatIn.nombre = catEg.nombre;
            OcatIn.largo = catEg.largo;
            cajDestino _desgiro = new cajDestino();
            _desgiro.codigo = giro.idcod; // tx_dat_giro.Text;
            _desgiro.nombre = giro.largo; // tx_ctaGiro.Text;
            _desgiro.largo = giro.ctades;  // eti_nomCtaGiro.Text;
            giro.idcod = ccdes;  // Ocajd.codigo;  // IDGiroConto de la cuenta egreso origen
            //string corre = correlativo(conn, ((tipcta == "cassaomg") ? "MCA" : "MCO"), int.Parse(fecOp.Substring(6,4)));
            string corre = correlativo(conn, "MCO", int.Parse(fecOp.Substring(6, 4)));  // 16/01/2025 GIROS SOLO COMO DESTINO CTAS PERSONALES
            Oingresos.creaIngreso("personal", fecOp, OcatIn, mone, monto, tcamb,
                _desgiro, descrip, corre, giro, aop);    // _desgiro, descrip, corre, giro, tx_anno.Text
            Oingresos.grabaIngreso(conn);
            giro.idcod = tx_dat_giro.Text;
            OcatIn.codigo = "";
            OcatIn.nombre = "";
            OcatIn.largo = "";
            _desgiro.codigo = "";
            _desgiro.nombre = "";
            _desgiro.largo = "";
        }

    }
}
