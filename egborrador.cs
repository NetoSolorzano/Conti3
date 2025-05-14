using ADGV;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Windows.Forms;
//using static Google.Protobuf.Collections.MapField<TKey, TValue>;

namespace Conti3
{
    public partial class egborrador : Form1
    {
        string nomform = "egborrador";
        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";port=" + login.port + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data +
            ";ConnectionLifeTime=" + login.ctl + ";";
        // datos de la grilla
        internal DataTable dt_grilla = new DataTable();                             // 
        DataTable dtpro = new DataTable();                                          // proveedores
        DataTable dtasoc = new DataTable();                                         // cuentas asociación
        List<string> lista_ = new List<string>();                                   // cuentas personales 
        List<string> lista_OMG = new List<string>();                                // cuentas OMG
        List<string> lista_CAM = new List<string>();                                // categorias egresos
        List<string> lista_prov = new List<string>();                               // proveedores
        //
        preliminares Opreli = new preliminares();
        publicoConf conf = new publicoConf();
        cajDestino Ocajd = new cajDestino();                                        // Objeto cada de destino - desde donde sale el dinero
        catEgresos OcatEg = new catEgresos();
        provees Oprove = new provees();                                             // Objeto nombre asignado
        montos Omonto = new montos();                                               // Objeto monto
        monedas Omone = new monedas();
        giroConto Ogiro = new giroConto();
        Egresos Oegresos = new Egresos();
        Finan_Egres oFEgres = new Finan_Egres();
        ccolores OColores = new ccolores();
        tipcamDia tcDia = new tipcamDia();
        string nomForm = "";
        int diasAtroya = 0;                                                         // dias atras hasta donde mostrará la grilla
        int limCols = 1;                                                            // limite de columnas que muestra la grilla
        string ususvalid = "";                                                      // usuarios que aprueban 
        string codDol = "MON002";
        string ctasAsoc = "";                                                       // cuentas "asociación"
        string codEur = "MON003";
        string codSol = "MON001";
        string col1rafila = "";                                                     // color html de la 1ra fila en ingresos

        public egborrador()
        {
            InitializeComponent();
            chk_giroC_CheckedChanged(null, null);   // 
            //chk_asoc_CheckedChanged(null, null);
        }
        private void egborrador_Load(object sender, EventArgs e)
        {
            CargaINI(this);                         // colorea los objetos graficos
            sololee("T");                           // T=todos los campos, "" ó "C" campos comunes
            jalainfo();                             // jala variables de tabla enlace
            initCampos();                           // pone maximos y upper case de campos texto
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                oFEgres.jalacolores(conn, OColores, nomForm);
                toolboton(conn);
                CargaFormatos(conn);                        // jala datos de combos y demas
            }
            oFEgres.colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);    // "#caf44d", "#d9f684", "#ecf8c8"
            tx_descrip.BackColor = ColorTranslator.FromHtml(OColores.Fondo_suave);
            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml(OColores.Fondo_boton_graba);   //  "#667d97"
            Bt_graba.Image = null;
            //rb_pers.Focus();
        }
        private void egborrador_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N) Bt_add.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E) Bt_edit.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.A) Bt_anul.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O) Bt_ver.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P) Bt_print.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.B) Bt_aprob.PerformClick();   // 13/05/2025
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            string para1 = "";
            string para2 = "";
            string para3 = "";
            string para4 = "";
            if (keyData == Keys.F1 && (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION"))
            {
                if (Tx_ctaDes.Focused == true)
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
                            Tx_ctaDes.Text = ayu2.ReturnValueA[2];      // [1]
                            eti_nomCaja.Text = ayu2.ReturnValueA[1];    // [2]
                            Ocajd.codigo = ayu2.ReturnValueA[0];
                            Ocajd.nombre = ayu2.ReturnValueA[2];
                            Ocajd.largo = ayu2.ReturnValueA[1];
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
                if (Tx_nomProv.Focused == true)
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
                            eti_nomprovee.Text = ayu2.ReturnValueA[1]; // 31/08/2024 ya no usamos
                            Oprove.codigo = tx_dat_provee.Text;
                            Oprove.nombre = Tx_nomProv.Text;
                            SendKeys.Send("{Tab}");
                        }
                    }
                }
                if (tx_ctaGiro.Focused == true)
                {
                    para1 = "personal";  // (rb_omg.Checked == true) ? "omg" : "personal";
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
                return true;    // indicate that you handled this keystroke
            }
            // Call the base class
            return base.ProcessCmdKey(ref msg, keyData);
        }    // F1 
        private void jalaoc()
        {
            tx_tipcam.Text = Opreli.TipCamb.ToString("#0.000");
            tx_anno.Text = DateTime.Parse(Opreli.FechOper).Year.ToString();
            tx_idOper.Text = Opreli.IdMovim;                     // Opreli.IdOper.ToString();
            selecFecha1.Value = DateTime.Parse(Opreli.FechOper);
            Tx_fecha.Text = Opreli.FechOper;
            Tx_catEgre.Text = Opreli.CatEgreso.largo;
            eti_nomCat.Text = Opreli.CatEgreso.nombre;
            Tx_ctaDes.Text = Opreli.CajaDes.largo;
            eti_nomCaja.Text = Opreli.CajaDes.nombre;
            tx_descrip.Text = Opreli.Descrip;
            Tx_nomProv.Text = Opreli.Proveedor.nombre;
            tx_dat_provee.Text = Opreli.Proveedor.codigo;
            eti_nomprovee.Text = Opreli.Proveedor.nombre;
            tx_ctaban.Text = Opreli.Proveedor.cuenta;
            Tx_rucprov.Text = Opreli.Proveedor.ruc;
            cmb_mon.SelectedValue = Opreli.Moneda.codigo;
            tx_montoS.Text = Opreli.Monto.monOrige.ToString("#0.00");
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
            chk_pagado.Checked = (Opreli.Pagado == 1) ? true : false;
        }                                                   // muestra en el formulario los objetos de la clase Egresos
        private void CargaFormatos(MySqlConnection conn)
        {
            // categoria egreso
            DataRow[] depar = Program.dt_definic.Select("idtabella='CAM' and numero=1");
            foreach (DataRow row in depar)
            {
                lista_CAM.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_catEgre.Values = lista_CAM.ToArray();

            // moneda por defecto
            depar = Program.dt_definic.Select("idtabella='MON' and idcodice='MON001'");
            foreach (DataRow row in depar)
            {
                Omone.codigo = row["idcodice"].ToString();    // la moneda
                Omone.nombre = row["descrizione"].ToString();    // en soles es
                Omone.siglas = row["descrizionerid"].ToString();    // por defecto
            }

            // cuenta personales destino
            depar = Program.dt_definic.Select("idtabella='CON' and numero=1");
            lista_.Clear();
            lista_ = new List<string>();
            foreach (DataRow row in depar)
            {
                lista_.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            tx_ctaGiro.Values = lista_.ToArray();
            // cuentas OMG destino
            depar = Program.dt_definic.Select("idtabella='DES' and numero=1");
            foreach (DataRow row in depar)
            {
                lista_OMG.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            //
            string consulta = "SELECT idanagrafica,trim(upper(ragionesociale)) AS nombre,RUC,cuenta FROM anag_for WHERE stato=1";
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
            // monedas
            depar = Program.dt_definic.Select("idtabella='MON' and numero=1");
            cmb_mon.DataSource = depar.CopyToDataTable();
            cmb_mon.DisplayMember = "descrizionerid";
            cmb_mon.ValueMember = "idcodice";
            // combo asociacion
            string consu = "select descrizione,idcodice,descrizionerid FROM desc_con WHERE numero=1 AND idcodice IN (" + ctasAsoc + ")";
            using (MySqlCommand micon = new MySqlCommand(consu, conn))
            {
                using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                {
                    da.Fill(dtasoc);
                }
            }
            cmb_asoc.DataSource = dtasoc;
            cmb_asoc.DisplayMember = "descrizione";
            cmb_asoc.ValueMember = "idcodice";
            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml("#f5510f");
            Bt_graba.Image = null;
        }
        private void jalainfo()
        {
            nomForm = this.Name;
            //DataRow[] row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='diasAtras'");
            //row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='limCols'");
            DataRow[] row = Program.dt_enlaces.Select("formulario='" + nomForm + "'");
            foreach (DataRow data in row)
            {
                if (data.ItemArray[2].ToString() == "grillas" && data.ItemArray[3].ToString() == "diasAtras")
                {
                    //diasAtroya = int.Parse(row[0]["valor"].ToString());
                    diasAtroya = int.Parse(data.ItemArray[5].ToString());
                }
                if (data.ItemArray[2].ToString() == "grillas" && data.ItemArray[3].ToString() == "limCols")
                {
                    //limCols = int.Parse(row[0]["valor"].ToString());
                    limCols = int.Parse(data.ItemArray[5].ToString());
                }
                if (data.ItemArray[2].ToString() == "grillas" && data.ItemArray[3].ToString() == "col1rafila")
                {
                    col1rafila = data.ItemArray[5].ToString();      // color html de la 1ra fila en ingresos
                }
                if (data.ItemArray[2].ToString() == "documento")
                {
                    if (data.ItemArray[3].ToString() == "usvalid") ususvalid = data.ItemArray[5].ToString();
                    if (data.ItemArray[3].ToString() == "ctasAsoc") ctasAsoc = data.ItemArray[5].ToString();
                }
            }
        }
        private void initCampos()
        {
            Bt_graba.Image = null;
            tx_anno.MaxLength = 4;
            tx_idOper.MaxLength = 15;
            Tx_catEgre.MaxLength = 50;
            Tx_catEgre.CharacterCasing = CharacterCasing.Upper;
            Tx_ctaDes.CharacterCasing = CharacterCasing.Upper;
            Tx_ctaDes.MaxLength = 50;
            Tx_nomProv.MaxLength = 50;
            tx_descrip.MaxLength = 93;  // 09/05/2025
            tx_ctaban.MaxLength = 20;   // cuenta bancaria 
            Tx_rucprov.MaxLength = 11;  // ruc
            //
            cmb_asoc.SelectedIndex = -1;
            cmb_mon.SelectedIndex = -1;
        }                                               // inicializa ancho de campos y upper case
        private void jala_ultimo(string _idop)
        {
            limpiaTE();
            limpiaObj("no");
            string fecOp = "";              // fecha de operacion
            decimal tipca = 0;              // tip cambio del monto origen
            string descr = "";              // descripcion de la operacion
            string idmov = "";              // id del movimiento
            string opera = "";              // operador, la que ó el que registra
            string aprob = "";              // usuario aprobador
            int pagad = 0;                  // 
            bool exito;
            // 1ro buscamos el registro si esta en la grilla
            DataRow[] fila = dt_grilla.Select("ID_MOVIM=" + oFEgres.CDerecha("00000" + _idop, 6));
            if (fila.Length > 0 && fila[0].ItemArray[0].ToString() != "")
            {
                exito = true;
                if (exito)
                {
                    fecOp = fila[0]["fecha"].ToString();    // advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);
                    OcatEg.codigo = fila[0]["IDCategoria"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["IDCategoria"].Value.ToString();
                    OcatEg.nombre = fila[0]["EGRESO"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["EGRESO"].Value.ToString();
                    OcatEg.largo = fila[0]["DET_EGRESO"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["DET_EGRESO"].Value.ToString();
                    Omone.codigo = fila[0]["codimon"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = fila[0]["MON"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["MON"].Value.ToString();
                    Omone.nombre = fila[0]["nombmon"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = fila[0]["codimon"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(fila[0]["MONTO"].ToString());   // decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["MONTO"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(fila[0]["TCMonOri"].ToString());   // decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(fila[0]["ImportoDU"].ToString());   // decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoDU"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(fila[0]["T_C"].ToString());   // decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["T_C"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(fila[0]["ImportoSU"].ToString());   // decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoSU"].Value.ToString());
                    tipca = decimal.Parse(fila[0]["TCMonOri"].ToString());   // decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Ocajd.codigo = fila[0]["IDConto"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["IDConto"].Value.ToString();
                    Ocajd.nombre = fila[0]["CUENTA"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["CUENTA"].Value.ToString();
                    Ocajd.largo = fila[0]["DET_CUENTA"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["DET_CUENTA"].Value.ToString();
                    Oprove.codigo = fila[0]["idanagrafica"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["idanagrafica"].Value.ToString();
                    Oprove.nombre = fila[0]["PROVEEDOR"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["PROVEEDOR"].Value.ToString();
                    Oprove.cuenta = fila[0]["cuentaB"].ToString();               // 
                    Oprove.ruc = fila[0]["RUC"].ToString();                  // 
                    descr = fila[0]["DESCRIPCION"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    idmov = fila[0]["ID_MOVIM"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString();
                    Ogiro.ctades = fila[0]["CTA_GIRO"].ToString();    // advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_GIRO"].Value.ToString();
                    Ogiro.tipodes = fila[0]["GIRO_CTA"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["GIRO_CTA"].Value.ToString();
                    Ogiro.codigo = fila[0]["CodGiro"].ToString();    // advancedDataGridView1.Rows[e.RowIndex].Cells["CodGiro"].Value.ToString();
                    Ogiro.largo = fila[0]["CTA_DESTINO"].ToString();     // advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_DESTINO"].Value.ToString();
                    Ogiro.idcod = fila[0]["IDGiroConto"].ToString();     // advancedDataGridView1.Rows[e.RowIndex].Cells["IDGiroConto"].Value.ToString();
                    opera = fila[0]["OPERADOR"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["OPERADOR"].Value.ToString();
                    aprob = fila[0]["APROBADOR"].ToString();   // advancedDataGridView1.Rows[e.RowIndex].Cells["APROBADOR"].Value.ToString();
                    pagad = int.Parse(fila[0]["pagado"].ToString());   // pagado 
                }
            }
            else
            {
                // 2do si no esta en la grilla, buscamos en la B.D.
                string[] retu = oFEgres.ValiIdOper((rb_omg.Checked == true) ? "PRE-OMG" : "PRE-PER", _idop, tx_anno.Text, "S");
                if (retu[0] == "")
                {
                    MessageBox.Show("No existe el código de operación");
                    exito = false;
                }
                else
                {
                    exito = true;
                    // ANNO,ID_MOVIM,FECHA,DET_CUENTA,EGRESO,MONEDA,MONTO,DESCRIPCION,TIP_CAMBIO,PROVEEDOR,OPERADOR,
                    //   0     1       2       3        4       5     6      7            8         9        10
                    // dia,APROBADOR,FEC_PROCESO,GIRO_CTA,IDGiroConto,ImportoDU,ImportoSU,idanagrafica,IDConto,IDCategoria,
                    //  11     12        13         14         15         16        17         18         19       20
                    // codimon,nombmon,TCMonOri,CUENTA,DET_EGRESO,CodGiro,CTA_DESTINO,CTA_GIRO,CASA,tipoE,RUC,cuentaB,pagado
                    //    21      22       23     24       25        26       27         28     29    30   31   32      33
                    fecOp = retu[2];   // fila[0]["fecha"].ToString();
                    OcatEg.codigo = retu[20];   // fila[0]["IDCategoria"].ToString();
                    OcatEg.nombre = retu[4];   // fila[0]["EGRESO"].ToString();
                    OcatEg.largo = retu[25];   // fila[0]["DET_EGRESO"].ToString();
                    Omone.codigo = retu[21];   // fila[0]["codimon"].ToString();
                    Omone.siglas = retu[5];   // fila[0]["MON"].ToString();
                    Omone.nombre = retu[22];   // fila[0]["nombmon"].ToString();
                    Omonto.codMOrige = retu[21];   // fila[0]["codimon"].ToString();
                    Omonto.monOrige = decimal.Parse(retu[6]);  // fila[0]["MONTO"].ToString()
                    Omonto.tipCOri = decimal.Parse(retu[23]);   // fila[0]["TCMonOri"].ToString()
                    Omonto.monDolar = decimal.Parse(retu[16]);  // fila[0]["ImportoDU"].ToString()
                    Omonto.tipCDol = decimal.Parse(retu[8]);   // fila[0]["T_C"].ToString()
                    Omonto.monSoles = decimal.Parse(retu[17]);  // fila[0]["ImportoSU"].ToString()
                    tipca = decimal.Parse(retu[23]);    // fila[0]["TCMonOri"].ToString()
                    Ocajd.codigo = retu[19];   // fila[0]["IDConto"].ToString();
                    Ocajd.nombre = retu[24];   // fila[0]["CUENTA"].ToString(); 
                    Ocajd.largo = retu[3];   // fila[0]["DET_CUENTA"].ToString();
                    Oprove.codigo = retu[18];   // fila[0]["idanagrafica"].ToString();
                    Oprove.nombre = retu[9];   // fila[0]["PROVEEDOR"].ToString();
                    Oprove.ruc = retu[31];              // ruc del proveedor
                    Oprove.cuenta = retu[32];           // cuenta bancaria del proveedor
                    descr = retu[7];   // fila[0]["DESCRIPCION"].ToString();
                    idmov = retu[1];   // fila[0]["ID_MOVIM"].ToString();
                    Ogiro.ctades = retu[28];   // fila[0]["CTA_GIRO"].ToString();
                    Ogiro.tipodes = retu[14];   // fila[0]["GIRO_CTA"].ToString();
                    Ogiro.codigo = retu[26];   // fila[0]["CodGiro"].ToString();
                    Ogiro.largo = retu[27];   // fila[0]["CTA_DESTINO"].ToString();
                    Ogiro.idcod = retu[15];   // fila[0]["IDGiroConto"].ToString();
                    opera = retu[10];   // fila[0]["OPERADOR"].ToString(); 
                    aprob = retu[12];   // fila[0]["APROBADOR"].ToString();
                    pagad = int.Parse(retu[33]);
                }
            }
            if (exito == true)
            {
                Opreli.creaPrelim(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, tipca, Ocajd, Oprove, descr, idmov, Ogiro, opera, "", pagad);
                // si ya fue aprobado el registro, ya no se puede editar o borrar 28/12/2024
                if ("EDICION,BORRAR,VALIDACION".Contains(Tx_modo.Text))
                {
                    if (aprob != "")
                    {
                        AutoClosingMessageBox.Show("Registro procesado, no es posible " + Tx_modo.Text, "Atención", 1500);
                        sololee("todos");
                    }
                    else
                    {
                        Bt_graba.Enabled = true;
                    }
                }
                else
                {
                    // nada
                }
                jalaoc();
            }
        }                                  // jala el registro 
        private void tipCambio(MySqlConnection condb)
        {
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
                return;
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
                            Omonto.tipCOri = Math.Round(dr.GetDecimal(1), 3);
                            tcDia.tcD = Omonto.tipCDol;
                            tcDia.tcE = Omonto.tipCOri;
                            if (Omonto.tipCDol <= 0 || Omonto.tipCOri <= 0)
                            {
                                MessageBox.Show("El tipo de cambio Dólares es: " + Omonto.tipCDol.ToString() + Environment.NewLine +
                                    "El tipo de cambio Euros es: " + Omonto.tipCOri.ToString(), "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                this.Close();
                            }
                        }
                    }
                    else
                    {
                        var aa = MessageBox.Show("No existen tipos de cambio para la fecha actual" + Environment.NewLine +
                            "Debe ingresarlos en este momento", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
        }
        private void actuaprov(MySqlConnection conn, string codigo, string Tx_rucprov, string tx_ctaban)
        {
            if (Tx_rucprov != "" || tx_ctaban != "")
            {
                string parte = "";
                if (Tx_rucprov != "") parte = "CodiceFiscale='" + Tx_rucprov + "'";
                if (tx_ctaban != "")
                {
                    if (parte == "") parte = "ContoCorrente='" + tx_ctaban + "'";
                    else parte = parte + ",ContoCorrente='" + tx_ctaban + "'";
                }
                string tarea = "update anagrafiche set " + parte + " where IDAnagrafica=@codi";
                using (MySqlCommand micon = new MySqlCommand(tarea, conn))
                {
                    micon.Parameters.AddWithValue("@codi", codigo);
                    micon.ExecuteNonQuery();
                }
            }
            Oprove.cuenta = tx_ctaban;
            Oprove.ruc = Tx_rucprov;
        }       // actualiza datos en proveedor
        private void actuapago()
        {
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Cells["pagado"].Value != null && row.Cells["pagado"].Value.ToString() == "1")
                {
                    row.Cells["Chk_PAG"].Value = 1;
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(col1rafila);   // 22/04/2025
                }
            }
        }                                               // marca check de PAGO en la grilla
        private void actuatabl(int idt, int vuc)
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                using (MySqlCommand micon = new MySqlCommand("update cassaprelim set pagado=@vuc where id=@idt", conn))
                {
                    micon.Parameters.AddWithValue("@idt", idt);
                    micon.Parameters.AddWithValue("@vuc", vuc);
                    micon.ExecuteNonQuery();
                }
            }
        }                               // actualiza marca en tabla

        #region marcación de checks de grilla
        private void marcaSelec(string modo)
        {
            rb_tos.Checked = false;
            rb_nin.Checked = false;
            switch (modo)
            {
                case "NUEVO":
                    pan_selM.Visible = false;
                    break;
                case "EDICION":
                    pan_selM.Visible = false;
                    //if (ususvalid.Contains(Program.vg_user) == true) pan_selM.Visible = true;
                    break;
                case "BORRAR":
                    pan_selM.Visible = false;
                    break;
                case "VISUALIZAR":
                    pan_selM.Visible = false;
                    break;
                case "VALIDACION":
                    if (ususvalid.Contains(Program.vg_user) == true) pan_selM.Visible = true;
                    break;
            }
        }
        private void marcaSelecGrilla(string modo)
        {
            for (int i=0; i < advancedDataGridView1.Rows.Count -1; i++)
            {
                DataGridViewRow row = advancedDataGridView1.Rows[i];
                if (modo == "marca")
                    row.Cells[0].Value = true;
                else row.Cells[0].Value = false;
            }
            /*  foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (modo == "marca")
                    row.Cells[0].Value = true;
                else row.Cells[0].Value = false;
            }   */
        }
        #endregion

        #region limpiadores, readonlys
        private void limpiaObj(string marca)
        {
            Ocajd.codigo = "";                                        // Objeto cada de destino - desde donde sale el dinero
            Ocajd.nombre = "";
            Ocajd.largo = "";
            OcatEg.codigo = "";
            OcatEg.nombre = "";
            OcatEg.largo = "";
            Oprove.codigo = "";
            Oprove.nombre = "";
            Omonto.codMOrige = "";                                    // Objeto monto
            Omonto.monDolar = 0;
            Omonto.monEuros = 0;
            Omonto.monOrige = 0;
            Omonto.monSoles = 0;
            Omonto.tipCDol = 0;
            Omonto.tipCOri = 0;
            Ogiro.codigo = "";
            Ogiro.ctades = "";
            Ogiro.idcod = "";
            Ogiro.largo = "";
            Ogiro.tipodes = "";
            if (marca == "todo")
            {
                advancedDataGridView1.DataSource = null;
                advancedDataGridView1.Columns.Clear();
                //advancedDataGridView1.Rows.Clear();
                //advancedDataGridView1
            }
        }
        private void limpiaTE() // limpia textbox, etiquetas, combos
        {
            tx_idOper.Clear();
            Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
            Tx_catEgre.Clear();
            eti_nomCat.Text = "";
            Tx_ctaDes.Clear();
            eti_nomCaja.Text = "";
            tx_descrip.Clear();
            Tx_nomProv.Clear();
            eti_nomprovee.Text = "";
            tx_dat_provee.Text = "";
            tx_montoS.Clear();
            tx_ctaban.Clear();
            Tx_rucprov.Clear();
            cmb_mon.SelectedIndex = -1; // no puede ser 0 porque el objeto moneda esta limpio 02/09/2024
            cmb_asoc.SelectedIndex = -1;    // 27-01-2025
            chk_giroC.Checked = false;
            chk_asoc.Checked = false;
            chk_pagado.Checked = false;
            tx_dat_asoc.Text = "";
        }
        private void escribe(string quien)  // pones los campos necesarios en readonly = false
        {
            tx_idOper.ReadOnly = false;
            tx_tipcam.ReadOnly = false;
            Tx_fecha.ReadOnly = false;
            selecFecha1.Enabled = true;
            Tx_catEgre.ReadOnly = false;
            Tx_ctaDes.ReadOnly = false;
            tx_descrip.ReadOnly = false;
            Tx_nomProv.ReadOnly = false;
            tx_montoS.ReadOnly = false;
            tx_ctaban.ReadOnly = false;
            Tx_rucprov.ReadOnly = false;
            cmb_mon.Enabled = true;
            Bt_graba.Enabled = true;
            chk_giroC.Enabled = true;
            chk_datSimil.Enabled = true;
            chk_pagado.Enabled = true;
        }
        private void sololee(string quien)  //    // T=todos los campos, "" ó "C" campos comunes
        {
            tx_idOper.ReadOnly = true;
            tx_tipcam.ReadOnly = true;
            selecFecha1.Enabled = false;
            Tx_fecha.ReadOnly = true;
            Tx_catEgre.ReadOnly = true;
            Tx_ctaDes.ReadOnly = true;
            tx_descrip.ReadOnly = true;
            Tx_nomProv.ReadOnly = true;
            tx_montoS.ReadOnly = true;
            tx_ctaban.ReadOnly = true;
            Tx_rucprov.ReadOnly = true;
            cmb_mon.Enabled = false;
            cmb_asoc.Enabled = false;
            Bt_graba.Enabled = false;
            chk_giroC.Enabled = false;
            chk_datSimil.Enabled = false;
            chk_pagado.Enabled = false;
        }
        #endregion

        #region datagridview
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            limpiaTE();
            limpiaObj("no");
            // CASA,ANNO,ID_MOVIM,FECHA,DET_CUENTA,EGRESO,MONEDA,MONTO,
            // DESCRIPCION,TIP_CAMBIO,PROVEEDOR,OPERADOR,dia,APROBADOR,FEC_PROCESO,
            // ImportoDU,ImportoSU,a.idanagrafica,a.IDConto,a.IDCategoria,
            // a.codimon,a.nombmon,a.TCMonOri,CUENTA,DET_EGRESO,a.CodGiro,
            // GIRO_CTA,a.IDGiroConto,CTA_DESTINO,CTA_GIRO,RUC,cuentaB,pagado
            if (e.RowIndex > -1)    // Tx_modo.Text != "NUEVO"
            {
                string fecOp = "";              // fecha de operacion
                decimal tipca = 0;              // tip cambio del monto origen
                string descr = "";              // descripcion de la operacion
                string idmov = "";              // id del movimiento
                string opera = "";              // operador, la que ó el que registra
                string aprob = "";              // usuario aprobador
                int pagad = 0;                  // marca de pagado 0=sin pago, 1=pagado o transferido
                if (true)
                {
                    if (advancedDataGridView1.Rows[e.RowIndex].Cells["tipoE"].Value.ToString() == "OMG")
                    {
                        rb_omg.Checked = true;
                    }
                    else
                    {
                        rb_pers.Checked = true;
                    }
                    fecOp = advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);
                    OcatEg.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDCategoria"].Value.ToString();
                    OcatEg.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["EGRESO"].Value.ToString();
                    OcatEg.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_EGRESO"].Value.ToString();
                    Omone.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = advancedDataGridView1.Rows[e.RowIndex].Cells["MON"].Value.ToString();
                    Omone.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["MONTO"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoDU"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["T_C"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImportoSU"].Value.ToString());
                    tipca = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TCMonOri"].Value.ToString());
                    Ocajd.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDConto"].Value.ToString();
                    Ocajd.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["CUENTA"].Value.ToString();
                    Ocajd.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_CUENTA"].Value.ToString();
                    Oprove.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["idanagrafica"].Value.ToString();
                    Oprove.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["PROVEEDOR"].Value.ToString();
                    Oprove.ruc = advancedDataGridView1.Rows[e.RowIndex].Cells["RUC"].Value.ToString();
                    Oprove.cuenta = advancedDataGridView1.Rows[e.RowIndex].Cells["cuentaB"].Value.ToString();
                    descr = advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    idmov = advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString();
                    Ogiro.ctades = advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_GIRO"].Value.ToString();
                    Ogiro.tipodes = advancedDataGridView1.Rows[e.RowIndex].Cells["GIRO_CTA"].Value.ToString();
                    Ogiro.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["CodGiro"].Value.ToString();
                    Ogiro.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["CTA_DESTINO"].Value.ToString();
                    Ogiro.idcod = advancedDataGridView1.Rows[e.RowIndex].Cells["IDGiroConto"].Value.ToString();
                    opera = advancedDataGridView1.Rows[e.RowIndex].Cells["OPERADOR"].Value.ToString();
                    aprob = advancedDataGridView1.Rows[e.RowIndex].Cells["APROBADOR"].Value.ToString();
                    pagad = int.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["pagado"].Value.ToString());
                }
                Opreli.creaPrelim(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, tipca, Ocajd, Oprove, descr, idmov, Ogiro, opera, "", pagad);
                // si ya fue aprobado el registro, ya no se puede editar 28/12/2024
                if ("EDICION,BORRAR,VALIDACION".Contains(Tx_modo.Text))
                {
                    if (aprob != "")
                    {
                        AutoClosingMessageBox.Show("Registro procesado, no es posible " + Tx_modo.Text, "Atención", 1500);
                        sololee("todos");
                    }
                    else
                    {
                        Bt_graba.Enabled = true;
                    }
                }
                else
                {
                    // nada
                }
                jalaoc();
            }
        }
        private void insFilaEnDataG(string _casa, string _corre)
        {
            // CASA,ANNO,ID_MOVIM,FECHA,DET_CUENTA,EGRESO,MONEDA,MONTO,
            // DESCRIPCION,TIP_CAMBIO,PROVEEDOR,OPERADOR,dia,APROBADOR,FEC_PROCESO,
            // ImportoDU,ImportoSU,a.idanagrafica,a.IDConto,a.IDCategoria,
            // a.codimon,a.nombmon,a.TCMonOri,CUENTA,DET_EGRESO,a.CodGiro,
            // GIRO_CTA,a.IDGiroConto,CTA_DESTINO,CTA_GIRO,RUC,cuentaB,pagado
            DataRow fila = dt_grilla.NewRow();
            string fecOp = selecFecha1.Value.Date.ToShortDateString();
            advancedDataGridView1.Rows[0].DefaultCellStyle.BackColor = System.Drawing.SystemColors.Window;  // 22/04/2025
            if (true)
            {
                fila["CASA"] = _casa;
                fila["ANNO"] = tx_anno.Text;
                fila["ID_MOVIM"] = _corre;
                fila["FECHA"] = fecOp;
                fila["DET_CUENTA"] = Ocajd.nombre;     // nombre cuenta destino
                fila["EGRESO"] = OcatEg.nombre;     // nombre categoria egreso
                fila["MON"] = Omone.siglas;      // siglas moneda origen
                fila["MONTO"] = Omonto.monOrige;    // valor origen
                fila["DESCRIPCION"] = tx_descrip.Text;
                fila["T_C"] = Omonto.tipCOri; // decimal.Parse(tx_tipcam.Text);
                fila["OPERADOR"] = Program.vg_user;
                fila["dia"] = DateTime.Now;
                // fila["APROBADOR"] = Program.vg_user;
                // fila["FEC_PROCESO"] = "";
                fila["ImportoDU"] = Omonto.monDolar;
                fila["ImportoSU"] = Omonto.monSoles;
                fila["idanagrafica"] = Oprove.codigo;
                fila["PROVEEDOR"] = Oprove.nombre;
                fila["cuentaB"] = Oprove.cuenta;
                fila["RUC"] = Oprove.ruc;
                fila["IDConto"] = Ocajd.codigo;
                fila["IDCategoria"] = OcatEg.codigo;
                fila["codimon"] = Omone.codigo;
                fila["nombmon"] = Omone.nombre;
                fila["TCMonOri"] = tx_tipcam.Text;
                fila["CUENTA"] = Ocajd.largo;
                fila["DET_EGRESO"] = OcatEg.largo;
                //
                fila["CodGiro"] = "";   // Ogiro.codigo;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["idgiroconto"] = Ogiro.idcod;
                fila["CTA_DESTINO"] = Ogiro.largo;
                fila["CTA_GIRO"] = Ogiro.ctades;
                fila["GIRO_CTA"] = Ogiro.tipodes;
                fila["tipoE"] = (rb_omg.Checked == true) ? "OMG" : "PER";
                fila["pagado"] = (chk_pagado.CheckState == CheckState.Checked) ? 1 : 0;
            }
            dt_grilla.Rows.InsertAt(fila, 0);
            advancedDataGridView1.CurrentCell = advancedDataGridView1.Rows[0].Cells[0];
            advancedDataGridView1.CurrentRow.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(col1rafila);   // 22/04/2025
            if ("NUEVO,EDICION".Contains(Tx_modo.Text))
            {
                // marcamos el CHK_PAG
                actuapago();
            }
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
                        string consulta = "ConPrel_cassa";
                        //if (rb_omg.Checked == true) consulta = "ConPrel_cassaOmg";        01/02/2025
                        //if (rb_pers.Checked == true) consulta = "ConPrel_cassaConti";     01/02/2025
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@Vdias", dAtras);
                            micon.Parameters.AddWithValue("@Vanno", 0);
                            micon.Parameters.AddWithValue("@Vidmov", "");
                            // muestra los registros no procesados
                            if ("NUEVO,VISUALIZAR".Contains(Tx_modo.Text))
                            {
                                micon.Parameters.AddWithValue("@proce", "T");  // T=todos los registros
                            }
                            // muestra todos los registros incluyendo los procesados
                            if ("EDICION,BORRAR,VALIDACION".Contains(Tx_modo.Text))
                            {
                                micon.Parameters.AddWithValue("@proce", "");  // en blanco=registros sin procesar
                            }
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_grilla.Clear();
                                dt_grilla.Columns.Clear();
                                da.Fill(dt_grilla);
                                advancedDataGridView1.DataSource = dt_grilla;
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
        private void armaGrilla(AdvancedDataGridView dgv_, int filasLim)
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
                if ("VALIDACION".Contains(Tx_modo.Text) && ususvalid.Contains(Program.vg_user)) // "EDICION,VALIDACION".Contains(Tx_modo.Text) &&
                {
                    DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                    chk.Name = "Chk_Val";
                    chk.HeaderText = "APROBADO";
                    //advancedDataGridView1.Columns.Add(chk);
                    advancedDataGridView1.Columns.Insert(0, chk);
                    advancedDataGridView1.ReadOnly = false;
                    for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
                    {
                        advancedDataGridView1.Columns[i].ReadOnly = true;    // acá ponemos todas las columnas en readonly menos la ultima con check
                    }
                    advancedDataGridView1.Columns["Chk_Val"].ReadOnly = false;

                    // si el registro ya fue aprobado, no debe dejar marcar el check 
                    for (int i = 0; i < advancedDataGridView1.Rows.Count - 1; i++)
                    {
                        if (advancedDataGridView1.Rows[i].Cells["APROBADOR"].Value.ToString() != "")   // si ya fue aprobado no debe dejar marcar
                        {
                            advancedDataGridView1.Rows[i].Cells["Chk_Val"].ReadOnly = true;
                            advancedDataGridView1.Rows[i].Cells["Chk_Val"].Value = 1;  // <-- no estoy seguro .. 
                        }
                    }
                }
                if ("NUEVO,EDICION".Contains(Tx_modo.Text) && ususvalid.Contains(Program.vg_user))
                {
                    DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                    chk.Name = "Chk_PAG";
                    chk.HeaderText = "PAGO";
                    chk.Width = 60;
                    advancedDataGridView1.Columns.Insert(0, chk);
                    advancedDataGridView1.ReadOnly = false;
                    for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
                    {
                        advancedDataGridView1.Columns[i].ReadOnly = true;    // acá ponemos todas las columnas en readonly menos la ultima con check
                    }
                    advancedDataGridView1.Columns["Chk_PAG"].ReadOnly = true;  // false 06/05/2025 no se puede cambiar 
                    // si el registro ya fue pagado, le pone el check
                    for (int i = 0; i < advancedDataGridView1.Rows.Count - 1; i++)
                    {
                        if (advancedDataGridView1.Rows[i].Cells["pagado"].Value.ToString() != "0")   // si ya fue aprobado no debe dejar marcar
                        {
                            advancedDataGridView1.Rows[i].Cells["Chk_PAG"].Value = 1;
                            advancedDataGridView1.Rows[i].DefaultCellStyle.BackColor = ColorTranslator.FromHtml(col1rafila);   // 22/04/2025
                        }
                    }
                }
            }
        }                 // ajusta el ancho de las columnas y muestra hasta el limite
        public void actFilaEnDataI(DataTable dt, string _casa, string _corre)
        {
            // CASA,ANNO,ID_MOVIM,FECHA,DET_CUENTA,EGRESO,MONEDA,MONTO,
            // DESCRIPCION,TIP_CAMBIO,PROVEEDOR,OPERADOR,dia,APROBADOR,FEC_PROCESO,
            // ImportoDU,ImportoSU,a.idanagrafica,a.IDConto,a.IDCategoria,
            // a.codimon,a.nombmon,a.TCMonOri,CUENTA,DET_EGRESO,a.CodGiro,
            // GIRO_CTA,a.IDGiroConto,CTA_DESTINO,CTA_GIRO,RUC,cuentaB,pagado
            string fecOp = selecFecha1.Value.Date.ToShortDateString();
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = dt.Rows[i];
                if (dr["ANNO"].ToString() == tx_anno.Text && dr["ID_MOVIM"].ToString() == oFEgres.CDerecha(_corre, 6))
                {
                    dr["CASA"] = _casa;
                    dr["ANNO"] = tx_anno.Text;
                    dr["ID_MOVIM"] = _corre;
                    dr["FECHA"] = fecOp;
                    dr["CUENTA"] = Ocajd.nombre;
                    dr["EGRESO"] = OcatEg.nombre;
                    dr["MON"] = Omone.siglas;
                    dr["MONTO"] = Omonto.monOrige;
                    dr["DESCRIPCION"] = tx_descrip.Text;
                    dr["T_C"] = Omonto.tipCOri; //  decimal.Parse(tx_tipcam.Text);
                    dr["OPERADOR"] = Opreli.Operador;
                    //dr["dia"] = ;
                    dr["APROBADOR"] = "";   // Opreli.Aprobador;    | el aprobador y fecha de proceso
                    //dr["FEC_PROCESO"] = ""; // Opreli.FechOper;     | no se editan, solo se muestran en modo VER
                    dr["ImportoDU"] = Omonto.monDolar;
                    dr["ImportoSU"] = Omonto.monSoles;
                    dr["idanagrafica"] = Oprove.codigo;
                    dr["PROVEEDOR"] = Oprove.nombre;
                    dr["cuentaB"] = Oprove.cuenta;
                    dr["RUC"] = Oprove.ruc;
                    dr["IDConto"] = Ocajd.codigo;
                    dr["IDCategoria"] = OcatEg.codigo;
                    dr["codimon"] = Omone.codigo;
                    dr["nombmon"] = Omone.nombre;
                    dr["TCMonOri"] = Omonto.tipCOri;
                    dr["DET_CUENTA"] = Ocajd.largo;
                    dr["DET_EGRESO"] = OcatEg.largo;
                    dr["CodGiro"] = "";   // Ogiro.codigo;
                    dr["GIRO_CTA"] = Ogiro.tipodes;
                    dr["idgiroconto"] = Ogiro.idcod;
                    dr["CTA_DESTINO"] = Ogiro.largo;
                    dr["CTA_GIRO"] = Ogiro.ctades;
                    dr["GIRO_CTA"] = Ogiro.tipodes;
                    dr["tipoE"] = (rb_omg.Checked == true) ? "OMG" : "PER";
                    dr["pagado"] = (chk_pagado.CheckState == CheckState.Checked) ? 1 : 0;
                }
                dr.AcceptChanges();
            }
            dt.AcceptChanges();
            if ("NUEVO,EDICION".Contains(Tx_modo.Text))
            {
                // marcamos el CHK_PAG 
                actuapago();
            }
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

        #region leaves y focus
        private void Tx_ctaDes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (Tx_ctaDes.Text.Trim() != "" && Tx_ctaDes.Text.Length >= 3)  // *************** 14/12/2024)
                    {
                        string[] vuelto = oFEgres.ValiCtaCon(Tx_ctaDes.Text, (rb_omg.Checked == true) ? "OMG" : "PER", "algo"); // Tx_ctaDes.Text, "OMG", "algo"
                        if (vuelto.Length > 0)
                        {
                            Ocajd.codigo = vuelto[0];
                            Ocajd.nombre = vuelto[1];
                            Ocajd.largo = vuelto[2];
                            eti_nomCaja.Text = Ocajd.nombre; //  = Ocajd.largo
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
        }
        private void tx_idOper_Validating(object sender, CancelEventArgs e)
        {
            if (tx_idOper.Text.Trim() != "" && Tx_modo.Text != "NUEVO")
            {
                jala_ultimo(tx_idOper.Text.Trim());
            }
        }     // busca en toda la base de datos
        private void tx_monto_Validating(object sender, CancelEventArgs e)
        {
            decimal monti = 0; decimal cambi = 0;
            decimal.TryParse(tx_montoS.Text, out monti);
            tx_montoS.Text = Math.Round(monti, 2).ToString("#,##0.00");
            decimal.TryParse(tx_tipcam.Text, out cambi);
            if (Tx_modo.Text == "NUEVO" && monti > 0)
            {
                Omonto.monOrige = monti;
                if (true)
                {
                    Omonto = oFEgres.calc_monedas(cmb_mon, monti, cambi);
                }

                if (Omone.codigo == codDol)
                {
                    Omonto.tipCDol = tcDia.tcD; // Omonto.tipCOri;
                    Omonto.tipCOri = tcDia.tcD;
                    Omonto.monEuros = 0;
                    Omonto.monDolar = decimal.Parse(tx_montoS.Text);
                    Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                    Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), decimal.Parse(tx_tipcam.Text));
                }
                if (Omone.codigo == codSol)
                {
                    Omonto.tipCDol = tcDia.tcD;
                    Omonto.tipCOri = tcDia.tcD;
                    Omonto.monEuros = 0;
                    Omonto.monSoles = decimal.Parse(tx_montoS.Text); // Omonto.monDolar * Omonto.tipCOri;
                    Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                    Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), Omonto.tipCDol);
                }
                if (Omone.codigo == codEur)
                {
                    Omonto.tipCDol = 0;
                    Omonto.tipCOri = tcDia.tcE;
                    Omonto.monEuros = decimal.Parse(tx_montoS.Text);
                    Omonto.monDolar = 0;
                    Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                    Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), Omonto.tipCOri);
                }
            }
        }
        private void tx_tipcam_Validating(object sender, CancelEventArgs e)
        {
            decimal monti = 0; decimal cambi = 0;
            decimal.TryParse(tx_montoS.Text, out monti);
            decimal.TryParse(tx_tipcam.Text, out cambi);
            tx_tipcam.Text = Math.Round(cambi, 3).ToString("#0.000");
            if (Tx_modo.Text == "NUEVO" && monti > 0)
            {
                Omonto.monOrige = monti;
                if (true)
                {
                    Omonto = oFEgres.calc_monedas(cmb_mon, monti, cambi);
                }
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
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                Tx_fecha.Text = selecFecha1.Value.Date.ToString("dd/MM/yyyy");
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
        private void Tx_ctaDes_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_ctaDes.Text.Trim() != "" && Tx_ctaDes.Text.Length >= 3)  // *************** 14/12/2024
                {
                    string[] vuelto = oFEgres.ValiCtaCon(Tx_ctaDes.Text, (rb_omg.Checked == true) ? "OMG":"PER", "algo");    // Tx_ctaDes.Text, "OMG", "algo"
                    if (vuelto.Length > 0 && vuelto[0] != "")
                    {
                        Ocajd.codigo = vuelto[0];
                        Ocajd.nombre = vuelto[1];
                        Ocajd.largo = vuelto[2];
                        eti_nomCaja.Text = Ocajd.nombre;
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
        private void Tx_ctaDes_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        }
        private void tx_descrip_Enter(object sender, EventArgs e)
        {
            //TextBox textBox = (TextBox)sender;    aca no porque al presionar enter
            //textBox.SelectAll();                  en la seleccion ... se borra 
        }
        private void rb_pers_CheckedChanged(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "" && rb_pers.Checked == true)
            {
                pan_p.Tag = "personal";
                //limpiaTE();
                //limpiaObj("todo");
                Tx_ctaDes.Text = "";
                eti_nomCaja.Text = "";
                Ocajd.codigo = ""; Ocajd.nombre = ""; Ocajd.largo = "";

                //jalaGrilla(diasAtroya, "");  // cassaconti    muestra datos de un dias atras hasta hoy
                Tx_ctaDes.Values = lista_.ToArray();
                //tx_ctaGiro.Values = lista_.ToArray();
                if (tx_tipcam.Text == "") tx_tipcam.Focus();
                else tx_idOper.Focus();
            }
        }
        private void rb_omg_CheckedChanged(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "" && rb_omg.Checked == true)
            {
                pan_p.Tag = "omg";
                //limpiaTE();
                //limpiaObj("todo");
                Tx_ctaDes.Text = "";
                eti_nomCaja.Text = "";
                Ocajd.codigo = ""; Ocajd.nombre = ""; Ocajd.largo = "";

                //jalaGrilla(diasAtroya, "");  // cassaomg     muestra datos de un dias atras hasta hoy
                Tx_ctaDes.Values = lista_OMG.ToArray();
                //tx_ctaGiro.Values = lista_.ToArray();   // tx_ctaGiro.Values = lista_OMG.ToArray();
                if (tx_tipcam.Text == "") tx_tipcam.Focus();
                else tx_idOper.Focus();
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
        private void Tx_catEgre_Enter(object sender, EventArgs e)
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
                    if (Tx_catEgre.Text.Trim() != "" && Tx_catEgre.Text.Length >= 3)  // *************** 14/12/2024)
                    {
                        if (Tx_catEgre.Text.Trim() != "")
                        {
                            DataRow[] nc = Program.dt_definic.Select("idtabella='CAM' and descrizione='" + Tx_catEgre.Text.Trim() + "'");
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
        private void Tx_nomProv_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_nomProv.Text.Trim().Length > 3 && (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION"))
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    // este segmento lo pase al leave()
                }
            }
        }
        private void Tx_nomProv_Leave(object sender, EventArgs e)
        {
            if (Tx_nomProv.Text.Trim() != "" && (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION"))
            {
                DataRow[] row;
                {
                    row = dtpro.Select("nombre='" + Tx_nomProv.Text.Trim() + "'");
                }
                if (row.Length > 0)
                {
                    tx_dat_provee.Text = row[0][0].ToString();     // idanagrafica
                    eti_nomprovee.Text = Tx_nomProv.Text.Trim(); // 31/08/2024 ya no usamos
                    tx_ctaban.Text = row[0][3].ToString();
                    Tx_rucprov.Text = row[0][2].ToString();
                    Oprove.nombre = Tx_nomProv.Text;
                    Oprove.codigo = tx_dat_provee.Text;
                    Oprove.ruc = Tx_rucprov.Text;
                    Oprove.cuenta = tx_ctaban.Text;
                }
                else
                {
                    Tx_nomProv.Clear();
                    tx_dat_provee.Clear();
                    Oprove.nombre = "";
                    Oprove.codigo = "";
                    Oprove.ruc = "";
                    Oprove.cuenta = "";
                    eti_nomprovee.Text = "";
                    tx_ctaban.Text = "";
                    Tx_rucprov.Text = "";
                    var aaa = MessageBox.Show("No existe el proveedor" + Environment.NewLine +
                        "Desea crearlo ahora?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aaa == DialogResult.Yes)
                    {
                        // llama a ayuda1
                        bt_Pnuevo_Click(null, null);
                    }
                }
            }
            /*
            if (Tx_modo.Text != "" && Tx_nomProv.Text.Trim() == "")
            {
                eti_nomprovee.Text = "";
                tx_dat_provee.Text = "";
                Tx_rucprov.Text = "";
                tx_ctaban.Text = "";
                Oprove.codigo = "";
                Oprove.nombre = "";
                Oprove.ruc = "";
                Oprove.cuenta = "";
            }
            if (Tx_modo.Text != "" && Tx_nomProv.Text.Trim() != "" && Oprove.codigo == "")
            {
                var aaa = MessageBox.Show("No existe el proveedor" + Environment.NewLine +
                                "Desea crearlo ahora?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    bt_Pnuevo_Click(null, null);
                }
                else
                {
                    Tx_nomProv.Text = "";
                    eti_nomprovee.Text = "";
                    tx_dat_provee.Text = "";
                    Tx_rucprov.Text = "";
                    tx_ctaban.Text = "";
                }
            }
            */
        }
        private void chk_giroC_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_giroC.CheckState == CheckState.Checked)
            {
                tx_ctaGiro.Visible = true;
                eti_nomCtaGiro.Visible = true;
                //
                chk_asoc.CheckState = CheckState.Unchecked;
            }
            else
            {
                tx_ctaGiro.Text = "";
                tx_ctaGiro.Visible = false;
                eti_nomCtaGiro.Text = "";
                eti_nomCtaGiro.Visible = false;
                tx_dat_giro.Text = "";
                Ogiro.codigo = "";
                Ogiro.ctades = "";
                Ogiro.largo = "";
                Ogiro.tipodes = "";
                Ogiro.idcod = "";
            }
        }
        private void chk_asoc_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_asoc.CheckState == CheckState.Checked)
            {
                chk_giroC.CheckState = CheckState.Unchecked;
                cmb_asoc.Enabled = true;
            }
            else
            {
                tx_dat_asoc.Text = "";
                cmb_asoc.SelectedIndex = -1;
                cmb_asoc.Enabled = false;
                Ogiro.codigo = "";
                Ogiro.ctades = "";
                Ogiro.largo = "";
                Ogiro.tipodes = "";
                Ogiro.idcod = "";
            }
        }
        private void chk_pagado_CheckedChanged(object sender, EventArgs e)
        {
            // hacemos algo de esto cuando grabamos
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
        private void yyy()
        {
            DataRow[] row;
            {
                row = Program.dt_definic.Select("idtabella='CON' and descrizione='" + tx_ctaGiro.Text.Trim() + "'");
            }
            if (row.Length > 0)
            {
                eti_nomCtaGiro.Text = row[0].ItemArray[3].ToString(); 
                //Ogiro.codigo = row[0].ItemArray[1].ToString();
                Ogiro.ctades = row[0].ItemArray[3].ToString();
                Ogiro.largo = row[0].ItemArray[2].ToString();
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
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void Tx_rucprov_Leave(object sender, EventArgs e)
        {
            if ("NUEVO,EDICION".Contains(Tx_modo.Text) && Tx_rucprov.Text != "")
            {
                Oprove.ruc = Tx_rucprov.Text;
            }
        }
        private void tx_ctaban_Leave(object sender, EventArgs e)
        {
            if ("NUEVO,EDICION".Contains(Tx_modo.Text) && tx_ctaban.Text != "")
            {
                Oprove.cuenta = tx_ctaban.Text;
            }
        }
        #endregion

        #region combos
        private void cmb_mon_SelectedValueChanged(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "" && (cmb_mon.SelectedValue != null && cmb_mon.SelectedValue.ToString() != ""))
            {
                Omone.codigo = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                Omone.siglas = cmb_mon.Text;    // siglas de la moneda
                Omonto.codMOrige = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                DataRow[] row = Program.dt_definic.Select("idtabella='MON' and idcodice='" + Omone.codigo + "'");
                Omone.nombre = row[0].ItemArray[2].ToString();
                if (tx_montoS.Text != "" && tx_tipcam.Text != "")
                {
                    Omonto.monOrige = decimal.Parse(tx_montoS.Text);
                    if (Omone.codigo == codDol)
                    {
                        Omonto.tipCDol = tcDia.tcD; // Omonto.tipCOri;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monDolar = decimal.Parse(tx_montoS.Text);
                        Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                        Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), decimal.Parse(tx_tipcam.Text));
                    }
                    if (Omone.codigo == codSol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monSoles = decimal.Parse(tx_montoS.Text); // Omonto.monDolar * Omonto.tipCOri;
                        Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                        Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), Omonto.tipCDol);
                    }
                    if (Omone.codigo == codEur)
                    {
                        Omonto.tipCDol = 0;
                        Omonto.tipCOri = tcDia.tcE;
                        Omonto.monEuros = decimal.Parse(tx_montoS.Text);
                        Omonto.monDolar = 0;
                        Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                        Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), Omonto.tipCOri);
                    }
                }
            }
        }   // selección de moneda
        private void cmb_mon_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_mon.SelectedIndex > -1 && (cmb_mon.SelectedValue != null && cmb_mon.SelectedValue.ToString() != ""))
            {
                Omone.codigo = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                Omone.siglas = cmb_mon.Text;    // siglas de la moneda
                Omonto.codMOrige = cmb_mon.SelectedValue.ToString();              // codigo de la moneda
                DataRow[] row = Program.dt_definic.Select("idtabella='MON' and idcodice='" + Omone.codigo + "'");
                if (row.Length > 1) Omone.nombre = row[0].ItemArray[2].ToString();
                if (tx_montoS.Text != "" && tx_tipcam.Text != "")
                {
                    Omonto.monOrige = decimal.Parse(tx_montoS.Text);
                    if (Omone.codigo == codDol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monDolar = decimal.Parse(tx_montoS.Text);
                        Omonto.monSoles = Omonto.monDolar * Omonto.tipCOri;
                        Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), decimal.Parse(tx_tipcam.Text));
                    }
                    if (Omone.codigo == codSol)
                    {
                        Omonto.tipCDol = tcDia.tcD;
                        Omonto.tipCOri = tcDia.tcD;
                        Omonto.monEuros = 0;
                        Omonto.monSoles = decimal.Parse(tx_montoS.Text); // Omonto.monDolar * Omonto.tipCOri;
                        Omonto.monDolar = Math.Round(Omonto.monSoles / Omonto.tipCDol, 2); // decimal.Parse(tx_monto.Text);
                        Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), Omonto.tipCDol);
                    }
                    if (Omone.codigo == codEur)
                    {
                        Omonto.tipCDol = 0;
                        Omonto.tipCOri = tcDia.tcE;
                        Omonto.monEuros = decimal.Parse(tx_montoS.Text);
                        Omonto.monDolar = 0;
                        Omonto.monSoles = Omonto.monEuros * Omonto.tipCOri;
                        Omonto = oFEgres.calc_monedas(cmb_mon, decimal.Parse(tx_montoS.Text), Omonto.tipCOri);
                    }
                }
            }
        }
        private void cmb_asoc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_asoc.SelectedIndex > -1 && Tx_modo.Text != "")
            {
                tx_dat_asoc.Text = cmb_asoc.SelectedValue.ToString();
                DataRow[] row = dtasoc.Select("idcodice='" + tx_dat_asoc.Text + "'");
                // generamos el giro
                Ogiro.largo = cmb_asoc.Text;
                Ogiro.codigo = "";    // PER<id_tabla>/OMG<id_tabla>
                Ogiro.ctades = row[0].ItemArray[2].ToString();    // nombre corto
                Ogiro.tipodes = (rb_omg.Checked == true) ? "OMG" : "PER";
                Ogiro.idcod = tx_dat_asoc.Text;     // idcodice de la cuenta
            }
        }
        #endregion

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
                if (Convert.ToString(row["btn7"]) == "S")
                {
                    this.Bt_aprob.Visible = true;
                }
                else { this.Bt_aprob.Visible = false; }
            }
            else
            {
                Bt_add.Visible = false;
                Bt_edit.Visible = false;
                Bt_anul.Visible = false;
                Bt_ver.Visible = false;
                Bt_print.Visible = false;
                Bt_aprob.Visible = false;
            }
        }
        private void Bt_add_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "NUEVO";
            limpiaObj("todo");
            limpiaTE();
            escribe("");
            rb_omg.Checked = false;
            rb_pers.Checked = false;
            selecFecha1.Enabled = true;
            Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
            tx_anno.Text = DateTime.Now.Date.Year.ToString();
            tx_anno.ReadOnly = true;
            tx_idOper.ReadOnly = true;
            chk_datSimil.Checked = false;
            chk_giroC.Enabled = true;
            jalaGrilla(diasAtroya, "");
            marcaSelec(Tx_modo.Text);
            rb_omg.Focus();
            if (true)   // tx_tipcam.Text == ""
            {
                tipCambio(null);
                //tx_tipcam.Focus();
            }
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "EDICION";
            limpiaObj("todo");
            limpiaTE();
            chk_datSimil.Checked = false;
            escribe("EDICION");    // sololee("")
            rb_omg.Checked = false;
            rb_pers.Checked = false;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;

            chk_giroC.Enabled = true;
            jalaGrilla(diasAtroya, "");
            marcaSelec(Tx_modo.Text);
            rb_omg.Focus();
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.delete40;
            Tx_modo.Text = "BORRAR";
            limpiaObj("todo");
            limpiaTE();
            sololee("");
            chk_datSimil.Checked = false;
            rb_omg.Checked = false;
            rb_pers.Checked = false;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;

            chk_giroC.Enabled = false;
            jalaGrilla(diasAtroya, "");
            marcaSelec(Tx_modo.Text);
            rb_omg.Focus();
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = null;
            Tx_modo.Text = "VISUALIZAR";
            limpiaObj("todo");
            limpiaTE();
            sololee("");
            rb_omg.Checked = false;
            rb_pers.Checked = false;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            jalaGrilla(diasAtroya, "");
            marcaSelec(Tx_modo.Text);
            rb_omg.Focus();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "IMPRIMIR";
        }
        private void Bt_aprob_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "VALIDACION";
            limpiaObj("todo");
            limpiaTE();
            chk_datSimil.Checked = false;
            escribe("EDICION");
            rb_omg.Checked = false;
            rb_pers.Checked = false;
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            jalaGrilla(diasAtroya, "");
            marcaSelec(Tx_modo.Text);
            //rb_omg.Focus();
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

        #region grabaciones botones
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
                        eti_nomprovee.Text = ayu1.ReturnValue1;
                        Tx_rucprov.Text = ayu1.ReturnValueA[2];
                        tx_ctaban.Text = ayu1.ReturnValueA[3];
                        Oprove.codigo = ayu1.ReturnValueA[0];
                        Oprove.nombre = ayu1.ReturnValueA[1];
                        Oprove.ruc = ayu1.ReturnValueA[2];
                        Oprove.cuenta = ayu1.ReturnValueA[3];
                        // idanagrafica,trim(upper(ragionesociale)) AS nombre,RUC,cuenta
                        DataRow dr = dtpro.NewRow();
                        dr[0] = Oprove.codigo;
                        dr[1] = Oprove.nombre.Replace("\r\n", string.Empty);
                        dr[2] = Oprove.ruc;
                        dr[3] = Oprove.cuenta;
                        dtpro.Rows.Add(dr);
                        dtpro.AcceptChanges();
                        lista_prov.Add(Oprove.nombre);
                        Tx_nomProv.Values = lista_prov.ToArray();
                        SendKeys.Send("{Tab}");
                    }
                }
            }
        }
        private void Bt_graba_Enter(object sender, EventArgs e)
        {
            Bt_graba.BackColor = Color.DarkSeaGreen;
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
            #region validaciones
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                // categoria de egreso no es obligatorio, no se valida en nuevo o edicion, solo en validación ... 27/01/2025
                if (tx_tipcam.Text == "")
                {
                    MessageBox.Show("Ingrese el tipo de cambio", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    tx_tipcam.Focus();
                    return;
                }
                if (Tx_ctaDes.Text == "")
                {
                    MessageBox.Show("Ingrese la cuenta de destino", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_ctaDes.Focus();
                    return;
                }
                if (Tx_ctaDes.Text.Trim() != Ocajd.largo)
                {
                    MessageBox.Show("Cuenta de destino incompleta!", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_ctaDes.Focus();
                    return;
                }
                if (cmb_mon.SelectedIndex < 0)
                {
                    MessageBox.Show("Seleccione la moneda", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    cmb_mon.Focus();
                    return;
                }
                if (tx_montoS.Text == "")
                {
                    MessageBox.Show("Ingrese el importe del egreso", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    tx_montoS.Focus();
                    return;
                }
                if (decimal.Parse(tx_montoS.Text) == 0)
                {
                    MessageBox.Show("Ingrese importe mayor a cero", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    tx_montoS.Focus();
                    return;
                }
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
                    if (eti_nomCtaGiro.Text.Trim() != Ogiro.ctades)
                    {
                        MessageBox.Show("Cuenta de Giro no completa", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        tx_ctaGiro.Select();
                        return;
                    }
                }
                // 28/04/2025  validaciones para proveedor
                if (Tx_nomProv.Text == "")
                {
                    MessageBox.Show("No tiene dato de proveedor" + Environment.NewLine +
                        "Debe ingresarlo", "Complete la información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    {
                        Tx_nomProv.Focus();
                        return;
                    }
                }
                // if (Tx_nomProv.Text != "" && (Tx_rucprov.Text.Trim() == "" || tx_ctaban.Text.Trim() == ""))
                if (Tx_nomProv.Text != "" && (tx_ctaban.Text.Trim() == ""))
                {
                    MessageBox.Show("Debe completar los datos del proveedor" + Environment.NewLine +
                        "falta la Cuenta bancaria", "Complete la información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    {
                        tx_ctaban.Focus();
                        return;
                    }
                }
            }
            #endregion

            if (Tx_modo.Text == "NUEVO")
            {
                var aa = MessageBox.Show("Desea grabar el registro actual?","Confirme por favor",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                if (aa == DialogResult.Yes)
                {
                    graba();
                    Tx_catEgre.Focus();
                }
                else { return; }
            }
            if (Tx_modo.Text == "EDICION")
            {
                if (tx_idOper.Text == "")
                {
                    MessageBox.Show("No hay registro que Editar!", "Identificador en blanco", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }
                var aa = MessageBox.Show("Desea modificar el registro actual?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aa == DialogResult.Yes)
                {
                    edita();
                    limpiaObj("no");    // limpia todo mejos grilla
                    limpiaTE();
                }
                else { return; }
            }
            if (Tx_modo.Text == "BORRAR")
            {
                // no se deben poder borrar registros que hayan sido aprobados y procesados
                // la validacion respectiva para no poder borrar registros procesados esta hecho en dobleclick de la grilla
                if (tx_idOper.Text == "")
                {
                    MessageBox.Show("No hay registro que borrar!", "Identificador en blanco", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }
                var aaa = MessageBox.Show("Confirma que desea BORRAR el registro?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    string tabla = "cassaprelim";
                    //dt_grilla.TableName = "dt_grillaE";
                    oFEgres.graba_borrar(tabla, tx_anno.Text, oFEgres.CDerecha("000000000000000" + tx_idOper.Text, 15), dt_grilla);
                    limpiaObj("no");
                    limpiaTE();
                }
            }
            if (Tx_modo.Text == "VALIDACION")
            {
                // buscamos en la grilla los checks que esten marcados y el campo aprobador=""
                // con esos registros marcados, fila por fila:
                //      1. creamos el objeto OEgresos y grabamos el objeto
                //      2. actualizamos la tabla preliminares con el aprobador y la fecha de proceso
                //      3. actualizamos la grilla, eliminamos el registro puesto que ya fue aprobado y no debe aparecer
                // al concluir avisamos y borramos todo
                int cta = 0; int cok = 0;
                using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                {
                    conn.Open();
                    foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                    {
                        if (row.Cells["ID_MOVIM"].Value != null)
                        {
                            if (row.Cells["Chk_Val"].Value != null && (bool)row.Cells["Chk_Val"].Value == true && row.Cells["APROBADOR"].Value.ToString() == "")
                            {
                                cta += 1;
                                if (!string.IsNullOrEmpty(row.Cells["IDCategoria"].Value.ToString()) &&
                                    !string.IsNullOrEmpty(row.Cells["codimon"].Value.ToString()) &&
                                    !string.IsNullOrEmpty(row.Cells["MONTO"].Value.ToString()) &&
                                    !string.IsNullOrEmpty(row.Cells["IDConto"].Value.ToString()))
                                {
                                    cok += 1;
                                    string fecOp = DateTime.Now.Date.ToShortDateString();    // debe ser la fecha del dia de la aprobacion
                                    string descr = row.Cells["DESCRIPCION"].Value.ToString();
                                    string corre = "000000000" + row.Cells["ID_MOVIM"].Value.ToString();    // 15 caract en total, 9 ceros + 6 caract movim
                                    OcatEg.largo = row.Cells["DET_EGRESO"].Value.ToString();
                                    OcatEg.nombre = row.Cells["EGRESO"].Value.ToString();
                                    OcatEg.codigo = row.Cells["IDCategoria"].Value.ToString();
                                    Omone.codigo = row.Cells["codimon"].Value.ToString();
                                    Omone.siglas = row.Cells["MON"].Value.ToString();
                                    Omone.nombre = row.Cells["nombmon"].Value.ToString();
                                    Omonto.tipCOri = decimal.Parse(row.Cells["T_C"].Value.ToString());
                                    //Omonto.monEuros = decimal.Parse(row.Cells[].Value.ToString());
                                    Omonto.tipCDol = decimal.Parse(row.Cells["T_C"].Value.ToString());
                                    Omonto.monDolar = decimal.Parse(row.Cells["ImportoDU"].Value.ToString());
                                    Omonto.monOrige = decimal.Parse(row.Cells["MONTO"].Value.ToString());
                                    Omonto.codMOrige = row.Cells["codimon"].Value.ToString();
                                    Omonto.monSoles = decimal.Parse(row.Cells["ImportoSU"].Value.ToString());
                                    Ocajd.codigo = row.Cells["IDConto"].Value.ToString();
                                    Ocajd.nombre = row.Cells["CUENTA"].Value.ToString();
                                    Ocajd.largo = row.Cells["DET_CUENTA"].Value.ToString();
                                    Oprove.codigo = row.Cells["IDAnagrafica"].Value.ToString();
                                    Oprove.nombre = row.Cells["PROVEEDOR"].Value.ToString();
                                    Ogiro.idcod = row.Cells["IDGiroConto"].Value.ToString();     // tx_dat_giro.Text;
                                    Ogiro.ctades = row.Cells["CTA_GIRO"].Value.ToString();      // eti_nomCtaGiro.Text;
                                    Ogiro.tipodes = row.Cells["GIRO_CTA"].Value.ToString();     // "PER";  // los giroconto siempre tienen como destino una cuenta personal ... 16/01/205
                                    Ogiro.largo = row.Cells["CTA_DESTINO"].Value.ToString();     // tx_ctaGiro.Text;
                                    // Ogiro.codigo = row.Cells[].Value.ToString();    // null;
                                    string yea = row.Cells["ANNO"].Value.ToString(); // DateTime.Now.Year.ToString();  // año real en que se genera el egreso
                                    string correE = oFEgres.correlativo(conn, (row.Cells["tipoE"].Value.ToString() == "OMG") ? "MCA" : "MCO", int.Parse(yea));
                                    // OJO que el tipo de cambio que se esta usando en la aprobación es de la fecha de creación !!!  
                                    int pagad = int.Parse(row.Cells["pagado"].Value.ToString());
                                    Oegresos.creaEgreso((row.Cells["tipoE"].Value.ToString() == "OMG") ? "omg" : "personal", fecOp, OcatEg, Omone, Omonto, decimal.Parse(row.Cells["T_C"].Value.ToString()),
                                                Ocajd, Oprove, descr, correE, Ogiro, yea);
                                    Oegresos.grabaEgreso(conn);
                                    // 
                                    Opreli.creaPrelim((row.Cells["tipoE"].Value.ToString() == "OMG") ? "omg" : "personal", row.Cells["FECHA"].Value.ToString(),
                                        OcatEg, Omone, Omonto, decimal.Parse(row.Cells["T_C"].Value.ToString()),
                                        Ocajd, Oprove, descr, corre, Ogiro, "", Program.vg_user, pagad);
                                    Opreli.actuaPrelim(conn, yea, corre);
                                    // si tiene giro, lo genera
                                    if (Ogiro.idcod != null && Ogiro.idcod != "")   // chk_giroC.CheckState == CheckState.Checked
                                    {
                                        //oFEgres.oper_giro(conn, Ogiro, (row.Cells["tipoE"].Value.ToString() == "OMG") ? "cassaomg" : "cassaconti", OcatEg, fecOp, Omone, Omonto, decimal.Parse(tx_tipcam.Text), descr, Ocajd.codigo);
                                        oFEgres.oper_giro(conn, Ogiro, (row.Cells["tipoE"].Value.ToString() == "OMG") ? "cassaomg" : "cassaconti", OcatEg, fecOp, Omone, Omonto, decimal.Parse(row.Cells["T_C"].Value.ToString()), descr, Ocajd.codigo, row.Cells["ANNO"].Value.ToString());
                                    }
                                }
                                Oegresos.limpia();
                                Opreli.limpia();
                                limpiaObj("");
                            }
                        }
                    }
                }
                if (cta > 0)
                {
                    if (cta == cok)
                    {
                        MessageBox.Show("Todo se procesó correctamente", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Registros incompletos no procesados", "Error en procesamiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    limpiaObj("todo");
                    limpiaTE();
                }
            }
        }
        private void graba()
        {
            string corre = "";
            string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
            decimal tipCam = 0;
            decimal.TryParse(tx_tipcam.Text, out tipCam);
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
                    // si tiene datos de proveedor actualiza (ruc y cuenta corriente) si tiene dato el campo
                    if (!string.IsNullOrEmpty(Oprove.codigo))
                    {
                        actuaprov(conn, Oprove.codigo.Trim(), Tx_rucprov.Text.Trim(), tx_ctaban.Text.Trim());
                    }
                    if (true)
                    {
                        corre = oFEgres.correlativo(conn, "MCP", int.Parse(tx_anno.Text));
                        if (corre != "error" && corre != "")
                        {
                            int pagad = (chk_pagado.CheckState == CheckState.Checked) ? 1 : 0; 
                            try
                            {
                                Opreli.creaPrelim(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                                        Ocajd, Oprove, tx_descrip.Text, corre, Ogiro, Program.vg_user, "", pagad);
                                Opreli.grabaPrelim(conn);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Error en grabar preliminar");
                                return;
                            }
                            insFilaEnDataG("LIM", oFEgres.CDerecha("00000" + corre, 6));       // inserta el registro nuevo en la grilla
                        }
                        else
                        {
                            MessageBox.Show("Error en grabar los datos del ingreso", "No se completo la operación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            if (chk_datSimil.CheckState == CheckState.Checked)
            {
                //jala_ultimo(oFEgres.CDerecha("00000" + corre, 6));
                Opreli.IdMovim = "";
                tx_idOper.Clear();
                jalaoc();   // pinta los datos en la pantalla
                chk_datSimil.Checked = true;
            }
            else
            {
                limpiaObj("no");
                limpiaTE();
            }
        }
        private void edita()
        {
            {
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
                            if (!string.IsNullOrEmpty(Oprove.codigo))
                            {
                                actuaprov(conn, Oprove.codigo.Trim(), Tx_rucprov.Text.Trim(), tx_ctaban.Text.Trim());
                            }
                            string fecOp = Tx_fecha.Text;    // selecFecha1.Value.Date.ToShortDateString();
                            decimal tipCam = 0;
                            decimal.TryParse(tx_tipcam.Text, out tipCam);
                            int pagad = (chk_pagado.CheckState == CheckState.Checked) ? 1 : 0;
                            string corre = oFEgres.CDerecha("000000000000000" + tx_idOper.Text, 15);
                            Opreli.creaPrelim(pan_p.Tag.ToString(), fecOp, OcatEg, Omone, Omonto, decimal.Parse(tx_tipcam.Text),
                                Ocajd, Oprove, tx_descrip.Text, corre, Ogiro, Program.vg_user, "", pagad);
                            Opreli.EditaPrelim(conn, tx_anno.Text, corre);
                            //
                            actFilaEnDataI(dt_grilla, "LIM", tx_idOper.Text);
                        }
                    }
                }
            }
        }
        #endregion

        private void egborrador_Click(object sender, EventArgs e)
        {
            this.Activate();
            this.BringToFront();
        }

    }
}
