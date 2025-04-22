using ADGV;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Conti3
{
    public partial class Finan_camion : Form1
    {
        string nomform = "Finan_camion";
        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";port=" + login.port + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data +
            ";ConnectionLifeTime=" + login.ctl + ";";
        // datos de la grilla
        internal DataTable dt_grilla = new DataTable();
        //
        publicoConf conf = new publicoConf();
        cajDestino Ocajd = new cajDestino();                                        // Objeto cada de destino - desde donde sale el dinero
        provees Oasig = new provees();                                                // Objeto nombre asignado
        montos Omonto = new montos();                                               // Objeto monto
        monedas Omone = new monedas();
        camiones Ocamion = new camiones();
        Finan_Egres oFEgres = new Finan_Egres();
        ccolores OColores = new ccolores();
        string nomForm = "";
        int diasAtroya = 0;                                                         // dias atras hasta donde mostrará la grilla
        int limCols = 1;                                                            // limite de columnas que muestra la grilla
        string col1rafila = "";                                                     // color html de la 1ra fila en ingresos

        public Finan_camion()
        {
            InitializeComponent();
            CargaINI(this);                         // colorea los objetos graficos
            CargaFormatos();                        // jala datos de combos y demas
            sololee("T");                           // T=todos los campos, "" ó "C" campos comunes
            jalainfo();                             // jala variables de tabla enlace
            initCampos();                           // pone maximos y upper case de campos texto
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                oFEgres.jalacolores(conn, OColores, nomForm);
                toolboton(conn);
            }
            oFEgres.colorea(this, OColores.Fondo_fuerte, OColores.Fondo_normal, OColores.Fondo_suave);    // "#caf44d", "#d9f684", "#ecf8c8"
            tx_descrip.BackColor = ColorTranslator.FromHtml(OColores.Fondo_suave);
            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml(OColores.Fondo_boton_graba);   //  "#667d97"
            Bt_graba.Image = null;
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
                if (Tx_asignado.Focused == true)
                {
                    para1 = "camion";       // : "personal";
                    para2 = "asignado";     // nombres asignados
                    para3 = "activos";      // todos | activos
                    ayuda2 ayu2 = new ayuda2(para1, para2, para3, para4);
                    var result = ayu2.ShowDialog();
                    if (result == DialogResult.Cancel)
                    {
                        if (!string.IsNullOrEmpty(ayu2.ReturnValue1))   // 0=codigo, 1=descripCorta, 2=descripLarga
                        {
                            tx_dat_asignado.Text = ayu2.ReturnValueA[0];
                            Tx_asignado.Text = ayu2.ReturnValueA[2];   // [1]
                        }
                    }
                }
                if (Tx_ctaDes.Focused == true)
                {
                    para1 = "omg";  // : "personal";
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
            Tx_asignado.Text = Ocamion.Placa.nombre;
            tx_anno.Text = Ocamion.AnnoOp;
            tx_idOper.Text = Ocamion.IdOper.ToString();
            selecFecha1.Value = DateTime.Parse(Ocamion.FechOper);
            Tx_fecha.Text = Ocamion.FechOper;
            Tx_ctaDes.Text = Ocamion.CajaDes.largo;    // .nombre
            eti_nomCaja.Text = Ocamion.CajaDes.nombre;   // .largo
            cmb_mon.SelectedValue = Ocamion.Moneda.codigo;
            tx_montoS.Text = Ocamion.TotalS.ToString("#0.00");
            tx_tipcam.Text = Ocamion.TipCamb.ToString("#0.000");
            tx_descrip.Text = Ocamion.Descrip;
            Tx_combus.Text = Ocamion.Combust.ToString("#0.00");
            Tx_viati.Text = Ocamion.Viaticos.ToString("#0.00");
            tx_dat_rptosS.Text = Ocamion.Respuest.ToString("#0.00");
            if (Ocamion.RptsDol > 0)
            {
                Tx_rptos.Text = (Ocamion.Respuest / Ocamion.TipCamb).ToString("#0.00");
                chk_dol.CheckState = CheckState.Checked;
            }
            else
            {
                Tx_rptos.Text = tx_dat_rptosS.Text;
            }
            Tx_impues.Text = Ocamion.Impuests.ToString("#0.00");
            Tx_honor.Text = Ocamion.Honorar.ToString("#0.00");
            Tx_varios.Text = Ocamion.Varios.ToString("#0.00");
            sumador();
        }                                                   // muestra en el formulario los objetos de la clase Egresos
        private void CargaFormatos()
        {
            // moneda por defecto
            DataRow[] depar = Program.dt_definic.Select("idtabella='MON' and idcodice='MON001'");
            foreach (DataRow row in depar)
            {
                Omone.codigo = row["idcodice"].ToString();    // la moneda
                Omone.nombre = row["descrizione"].ToString();    // en soles es
                Omone.siglas = row["descrizionerid"].ToString();    // por defecto
            }

            // asignado/encargado almacen/camion
            depar = Program.dt_definic.Select("idtabella='CMN' and numero=1");
            List<string> lista_ = new List<string>();
            foreach (DataRow row in depar)
            {
                // lista_.Add(row["descrizionerid"].ToString().Trim().ToUpper());
                lista_.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_asignado.Values = lista_.ToArray();

            // cuenta destino
            depar = Program.dt_definic.Select("idtabella='DES' and numero=1");
            lista_.Clear();
            lista_ = new List<string>();
            foreach (DataRow row in depar)
            {
                // lista_.Add(row["descrizionerid"].ToString().Trim().ToUpper());
                lista_.Add(row["descrizione"].ToString().Trim().ToUpper());
            }
            Tx_ctaDes.Values = lista_.ToArray();
            
            // monedas
            depar = Program.dt_definic.Select("idtabella='MON' and numero=1");
            cmb_mon.DataSource = depar.CopyToDataTable();
            cmb_mon.DisplayMember = "descrizionerid";
            cmb_mon.ValueMember = "idcodice";

            // color de boton Bt_graba
            Bt_graba.BackColor = ColorTranslator.FromHtml("#f5510f");   // , "#e76433"
            Bt_graba.Image = null;
        }
        private void jalainfo()
        {
            nomForm = this.Name;
            DataRow[] row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='diasAtras'");
            diasAtroya = int.Parse(row[0]["valor"].ToString());
            row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='limCols'");
            limCols = int.Parse(row[0]["valor"].ToString());
            row = Program.dt_enlaces.Select("formulario='" + nomForm + "' and campo='grillas' and param='col1rafila'");
            col1rafila = row[0]["valor"].ToString();              // color html de la 1ra fila en ingresos
        }
        private void initCampos()
        {
            Bt_graba.Image = null;
            tx_anno.MaxLength = 4;
            Tx_asignado.MaxLength = 50;  // = 7
            Tx_asignado.CharacterCasing = CharacterCasing.Upper;
            tx_idOper.MaxLength = 15;
            Tx_ctaDes.CharacterCasing = CharacterCasing.Upper;
            Tx_ctaDes.MaxLength = 50;   // 50
            tx_descrip.MaxLength = 100;
        }                                               // inicializa ancho de campos y upper case
        private void sumador()
        {
            decimal tipCam = 0;
            decimal valCom = 0;
            decimal valVia = 0;
            decimal valRep = 0;
            decimal valImp = 0;
            decimal valHon = 0;
            decimal valVar = 0;
            decimal.TryParse(tx_tipcam.Text, out tipCam);
            decimal.TryParse(Tx_combus.Text, out valCom);
            decimal.TryParse(Tx_viati.Text, out valVia);
            decimal.TryParse(tx_dat_rptosS.Text, out valRep);   // valor en soles de repuestos
            decimal.TryParse(Tx_impues.Text, out valImp);
            decimal.TryParse(Tx_honor.Text, out valHon);
            decimal.TryParse(Tx_varios.Text, out valVar);
            decimal totSoles = valCom + valVia + valRep + valImp + valHon + valVar;
            decimal totDolor = (tipCam > 0) ? totSoles / tipCam : 0;
            tx_montoS.Text = totSoles.ToString("#,##0.00");
            Tx_totalR.Text = totSoles.ToString("#,##0.00");
            tx_montoD.Text = totDolor.ToString("#,##0.00");
            //
            Omonto.codMOrige = "";
            Omonto.monDolar = totDolor;
            Omonto.monOrige = totSoles;
            Omonto.monSoles = totSoles;
            Omonto.tipCDol = tipCam;
            Omonto.tipCOri = tipCam;
            //
        }                                                  // totaliza en soles y dolares
        private void jala_ultimo(string _idop)
        {
            string[] retu = oFEgres.ValiIdOper("CAMION", _idop, tx_anno.Text, "");
            if (retu[0] == "")
            {
                limpiaObj();
                limpiaTE();
                MessageBox.Show("No existe el código de operación");
            }
            else
            {
                string anOp = "";
                string fecOp = "";              // fecha de operacion
                decimal tipca = 0;              // tip cambio del monto origen
                string descr = "";              // descripcion de la operacion
                string idmov = "";              // id del movimiento
                decimal combust = 0;
                decimal viaticos = 0;
                decimal respuest = 0;
                decimal varios = 0;
                decimal honorar = 0;
                decimal impuests = 0;
                decimal totalS = 0;
                decimal totalD = 0;
                chk_dol.CheckState = CheckState.Unchecked;
                chk_datSimil.CheckState = CheckState.Unchecked;
                tx_dat_asignado.Text = "";
                tx_dat_rptosS.Text = "";
                int repsEnDol = 0;  // 0=repuestos es soles, 1=valor de repuestos en dolares
                if (true)
                {
                    // ANNO,ID_MOVIM,FECHA,CAMION,DESTINO,TOTAL_SOL,DESCRIPCION,usuario,CASA,
                    //   0     1       2      3      4       5           6         7      8  
                    // TOTAL_DOL,TIP_CAMBIO,dia,codimon,nombmon,MONEDA,IDDestino,DET_DESTINO,
                    //     9         10      11    12      13     14      15         16   
                    // ImpCarbS,ImpViaS,ImpRicS,ImpVariS,imphons,ImpImpS,repdol,IDCamion
                    //    17      18      19       20       21      22     23      24
                    anOp = retu[0].ToString();
                    fecOp = retu[2].ToString();
                    Oasig.codigo = retu[24].ToString();
                    Oasig.nombre = retu[3].ToString();
                    Omone.codigo = retu[12].ToString();
                    Omone.siglas = retu[14].ToString();
                    Omone.nombre = retu[13].ToString();
                    Omonto.codMOrige = retu[12].ToString();
                    Omonto.monOrige = decimal.Parse(retu[5].ToString());
                    Omonto.tipCOri = decimal.Parse(retu[10].ToString());
                    Omonto.monDolar = decimal.Parse(retu[9].ToString());
                    Omonto.tipCDol = decimal.Parse(retu[10].ToString());
                    Omonto.monSoles = decimal.Parse(retu[5].ToString());
                    tipca = decimal.Parse(retu[10].ToString());
                    Ocajd.codigo = retu[15].ToString();
                    Ocajd.nombre = retu[4].ToString();
                    Ocajd.largo = retu[16].ToString();
                    descr = retu[6];
                    idmov = retu[1];
                    combust = decimal.Parse(retu[17]);
                    viaticos = decimal.Parse(retu[18]);
                    respuest = decimal.Parse(retu[19]);
                    varios = decimal.Parse(retu[20]);
                    honorar = decimal.Parse(retu[21]);
                    impuests = decimal.Parse(retu[22]);
                    totalS = decimal.Parse(retu[5]);
                    totalD = decimal.Parse(retu[9]);
                    if (retu[23] == "1") repsEnDol = 1;
                }
                Ocamion.creaCamion(Oasig, idmov, fecOp, Ocajd, "nada", "nada",
                    Omone, tipca, descr, combust, viaticos, respuest, impuests,
                    honorar, varios, totalS, totalD, repsEnDol, anOp);
                jalaoc();
            }
        }                                  // jala el ultimo registro ingresado

        #region limpiadores, readonlys
        private void limpiaObj()
        {
            Ocajd.codigo = "";                                        // Objeto cada de destino - desde donde sale el dinero
            Ocajd.nombre = "";
            Ocajd.largo = "";
            Omonto.codMOrige = "";                                    // Objeto monto
            Omonto.monDolar = 0;
            Omonto.monEuros = 0;
            Omonto.monOrige = 0;
            Omonto.monSoles = 0;
            Omonto.tipCDol = 0;
            Omonto.tipCOri = 0;
            Oasig.codigo = "";
            Oasig.nombre = "";
        }
        private void limpiaTE() // limpia textbox, etiquetas, combos
        {
            var colno = Tx_varios.ForeColor;
            tx_idOper.Clear();
            //tx_tipcam.Clear();
            tx_anno.Text = "";
            Tx_asignado.Clear();
            tx_dat_asignado.Clear();
            Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
            Tx_ctaDes.Clear();
            eti_nomCaja.Text = "";
            tx_descrip.Clear();
            tx_montoS.Clear();
            tx_montoD.Clear();
            Tx_combus.Text = 0.ToString("#0.00");
            Tx_viati.Text = 0.ToString("#0.00");
            Tx_rptos.Text = 0.ToString("#0.00");
            Tx_rptos.Font = new Font("Verdana", 8, FontStyle.Regular); // este es el font normal
            Tx_rptos.ForeColor = colno;                                // con el color normal
            Tx_varios.Text = 0.ToString("#0.00");
            Tx_honor.Text = 0.ToString("#0.00");
            Tx_impues.Text = 0.ToString("#0.00");
            Tx_totalR.Text = 0.ToString("#0.00");
            cmb_mon.SelectedIndex = -1; // no puede ser 0 porque el objeto moneda esta limpio 02/09/2024
            chk_dol.CheckState = CheckState.Unchecked;
        }
        private void escribe(string quien)  // pones los campos necesarios en readonly = false
        {
            tx_idOper.ReadOnly = false;
            tx_tipcam.ReadOnly = false;
            Tx_asignado.ReadOnly = false;
            tx_dat_asignado.ReadOnly = true;
            tx_anno.ReadOnly = false;
            Tx_fecha.ReadOnly = false;
            Tx_ctaDes.ReadOnly = false;
            tx_descrip.ReadOnly = false;
            tx_montoS.ReadOnly = true;
            tx_montoD.ReadOnly = true;
            Tx_combus.ReadOnly = false;
            Tx_viati.ReadOnly = false;
            Tx_rptos.ReadOnly = false;
            Tx_varios.ReadOnly = false;
            Tx_honor.ReadOnly = false;
            Tx_impues.ReadOnly = false;
            Tx_totalR.ReadOnly = true;
            chk_dol.Enabled = true;
        }
        private void sololee(string quien)  //    // T=todos los campos, "" ó "C" campos comunes
        {
            tx_idOper.ReadOnly = true;
            tx_tipcam.ReadOnly = true;
            Tx_asignado.ReadOnly = true;
            tx_dat_asignado.ReadOnly = true;
            tx_anno.ReadOnly = true;
            Tx_fecha.ReadOnly = true;
            Tx_ctaDes.ReadOnly = true;
            tx_descrip.ReadOnly = true;
            tx_montoS.ReadOnly = true;
            tx_montoD.ReadOnly = true;
            Tx_combus.ReadOnly = true;
            Tx_viati.ReadOnly = true;
            Tx_rptos.ReadOnly = true;
            Tx_varios.ReadOnly = true;
            Tx_honor.ReadOnly = true;
            Tx_impues.ReadOnly = true;
            Tx_totalR.ReadOnly = true;
            chk_dol.Enabled = false;
        }
        #endregion

        #region datagridview
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            limpiaTE();
            limpiaObj();
            // CASA,AÑO,ID_MOVIM,FECHA,CAMION,DESTINO,TOTAL_SOL,TOTAL_DOL,TIP_CAMBIO,DESCRIPCION,
            // usuario,dia,codimon,nombmon,MONEDA,IDDestino,DET_DESTINO,
            // a.ImpCarbS,a.ImpViaS,a.ImpRicS,a.ImpVariS,a.imphons,a.ImpImpS,a.repdol,IDCamion 
            if (true)    // 21/02/2025  Tx_modo.Text != "NUEVO"
            {
                string anOp = "";
                string fecOp = "";              // fecha de operacion
                decimal tipca = 0;              // tip cambio del monto origen
                string descr = "";              // descripcion de la operacion
                string idmov = "";              // id del movimiento
                decimal combust = 0;
                decimal viaticos = 0;
                decimal respuest = 0;
                decimal varios = 0;
                decimal honorar = 0;
                decimal impuests = 0;
                decimal totalS = 0;
                decimal totalD = 0;
                chk_dol.CheckState = CheckState.Unchecked;
                chk_datSimil.CheckState = CheckState.Unchecked;
                tx_dat_asignado.Text = "";
                tx_dat_rptosS.Text = "";
                int repsEnDol = 0;  // 0=repuestos es soles, 1=valor de repuestos en dolares
                if (true)
                {
                    anOp = advancedDataGridView1.Rows[e.RowIndex].Cells["ANNO"].Value.ToString();
                    fecOp = advancedDataGridView1.Rows[e.RowIndex].Cells["FECHA"].Value.ToString().Substring(0, 10);
                    Oasig.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDCamion"].Value.ToString();
                    Oasig.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["CAMION"].Value.ToString();
                    Omone.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omone.siglas = advancedDataGridView1.Rows[e.RowIndex].Cells["MONEDA"].Value.ToString();
                    Omone.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["nombmon"].Value.ToString();
                    Omonto.codMOrige = advancedDataGridView1.Rows[e.RowIndex].Cells["codimon"].Value.ToString();
                    Omonto.monOrige = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TOTAL_SOL"].Value.ToString());
                    Omonto.tipCOri = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Omonto.monDolar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TOTAL_DOL"].Value.ToString());
                    Omonto.tipCDol = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Omonto.monSoles = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TOTAL_SOL"].Value.ToString());
                    tipca = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TIP_CAMBIO"].Value.ToString());
                    Ocajd.codigo = advancedDataGridView1.Rows[e.RowIndex].Cells["IDDestino"].Value.ToString();
                    Ocajd.nombre = advancedDataGridView1.Rows[e.RowIndex].Cells["DESTINO"].Value.ToString();
                    Ocajd.largo = advancedDataGridView1.Rows[e.RowIndex].Cells["DET_DESTINO"].Value.ToString();
                    descr = advancedDataGridView1.Rows[e.RowIndex].Cells["DESCRIPCION"].Value.ToString();
                    idmov = advancedDataGridView1.Rows[e.RowIndex].Cells["ID_MOVIM"].Value.ToString();
                    combust = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImpCarbS"].Value.ToString()); 
                    viaticos = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImpViaS"].Value.ToString());
                    respuest = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImpRicS"].Value.ToString());
                    varios = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImpVariS"].Value.ToString());
                    honorar = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["imphons"].Value.ToString());
                    impuests = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["ImpImpS"].Value.ToString());
                    totalS = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TOTAL_SOL"].Value.ToString());
                    totalD = decimal.Parse(advancedDataGridView1.Rows[e.RowIndex].Cells["TOTAL_DOL"].Value.ToString());
                    if (advancedDataGridView1.Rows[e.RowIndex].Cells["repdol"].Value.ToString() == "1") repsEnDol = 1;
                }
                sumador();
                Ocamion.creaCamion(Oasig, idmov, fecOp, Ocajd, "nada", "nada",
                    Omone, tipca, descr, combust, viaticos, respuest, impuests,
                    honorar, varios, totalS, totalD, repsEnDol, anOp);   //varios, 
                jalaoc();
            }
        }
        private void insFilaEnDataG(string _casa, string _corre)
        {
            // CASA,AÑO,ID_MOVIM,FECHA,CAMION,DESTINO,TOTAL_SOL,TOTAL_DOL,TIP_CAMBIO,DESCRIPCION,
            // usuario,dia,codimon,nombmon,MONEDA,IDDestino,DET_DESTINO

            DataRow fila = dt_grilla.NewRow();
            string fecOp = Tx_fecha.Text;
            advancedDataGridView1.Rows[0].DefaultCellStyle.BackColor = System.Drawing.SystemColors.Window;  // 21/04/2025
            if (true)
            {
                fila["CASA"] = _casa;
                fila["ANNO"] = tx_anno.Text;
                fila["ID_MOVIM"] = _corre;
                fila["FECHA"] = fecOp;
                fila["CAMION"] = Oasig.nombre;
                fila["DESTINO"] = Ocajd.nombre;     // nombre cuenta destino
                fila["TOTAL_SOL"] = Omonto.monSoles;
                fila["TOTAL_DOL"] = Omonto.monDolar;
                fila["TIP_CAMBIO"] = Omonto.tipCOri;
                fila["DESCRIPCION"] = tx_descrip.Text;
                fila["usuario"] = Program.vg_user;
                fila["codimon"] = Omonto.codMOrige;
                fila["nombmon"] = Omone.nombre;
                fila["MONEDA"] = Omone.siglas;      // siglas moneda origen
                fila["IDDestino"] = Ocajd.codigo;
                fila["DET_DESTINO"] = Ocajd.largo;
                fila["ImpCarbS"] = Tx_combus.Text;
                fila["ImpViaS"] = Tx_viati.Text;
                fila["ImpRicS"] = tx_dat_rptosS.Text; // (chk_dol.Checked == true) ? Tx_rptos.Text : 
                fila["ImpVariS"] = Tx_varios.Text;
                fila["imphons"] = Tx_honor.Text;
                fila["ImpImpS"] = Tx_impues.Text;
                fila["repdol"] = (chk_dol.Checked == true) ? "1" : "0";
                fila["IDCamion"] = Oasig.codigo;
            }
            dt_grilla.Rows.InsertAt(fila, 0);
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
                        string consulta = "ConCamion";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.CommandType = CommandType.StoredProcedure;
                            micon.Parameters.AddWithValue("@Vdias", dAtras);
                            micon.Parameters.AddWithValue("@Vanno", 0);
                            micon.Parameters.AddWithValue("@Vidmov", "");
                            using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                            {
                                dt_grilla.Clear();
                                dt_grilla.Columns.Clear();
                                da.Fill(dt_grilla);
                                advancedDataGridView1.DataSource = dt_grilla;
                            }
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
                                        if (Omonto.tipCDol <= 0 || Omonto.tipCOri <= 0)
                                        {
                                            MessageBox.Show("El tipo de cambio Dólares es: " + Omonto.tipCDol.ToString() + Environment.NewLine +
                                                "El tipo de cambio Euros es: " + Omonto.tipCOri.ToString(), "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        }
                                    }
                                }
                                else
                                {
                                    var aa = MessageBox.Show("No existen tipos de cambio para la fecha actual" + Environment.NewLine +
                                        "Desea ingresarlos en este momento?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (aa == DialogResult.Yes)
                                    {
                                        // llamada a formulario de tipos de cambio
                                    }
                                }
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
            string fecOp = Tx_fecha.Text;     // selecFecha1.Value.Date.ToShortDateString();
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = dt.Rows[i];
                if (dr["ANNO"].ToString() == tx_anno.Text && dr["ID_MOVIM"].ToString() == oFEgres.CDerecha(_corre, 6))
                {
                    if (true)
                    {
                        dr["CASA"] = _casa;
                        dr["ANNO"] = tx_anno.Text;
                        dr["ID_MOVIM"] = _corre;
                        dr["FECHA"] = fecOp;
                        dr["CAMION"] = Oasig.nombre;
                        dr["DESTINO"] = Ocajd.nombre;     // nombre cuenta destino
                        dr["TOTAL_SOL"] = Omonto.monSoles;
                        dr["TOTAL_DOL"] = Omonto.monDolar;
                        dr["TIP_CAMBIO"] = Omonto.tipCOri;
                        dr["DESCRIPCION"] = tx_descrip.Text;
                        dr["usuario"] = Program.vg_user;
                        dr["codimon"] = Omonto.codMOrige;
                        dr["nombmon"] = Omone.nombre;
                        dr["MONEDA"] = Omone.siglas;      // siglas moneda origen
                        dr["IDDestino"] = Ocajd.codigo;
                        dr["DET_DESTINO"] = Ocajd.largo;
                        dr["ImpCarbS"] = Tx_combus.Text;
                        dr["ImpViaS"] = Tx_viati.Text;
                        dr["ImpRicS"] = tx_dat_rptosS.Text; // (chk_dol.Checked == true) ? Tx_rptos.Text : 
                        dr["ImpVariS"] = Tx_varios.Text;
                        dr["imphons"] = Tx_honor.Text;
                        dr["ImpImpS"] = Tx_impues.Text;
                        dr["repdol"] = (chk_dol.Checked == true) ? "1" : "0";
                        dr["IDCamion"] = Oasig.codigo;
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

        #region leaves y focus
        private void Tx_ctaDes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (e.KeyChar == (char)13 || e.KeyChar == (char)09)
                {
                    if (Tx_ctaDes.Text.Trim() != "" && Tx_ctaDes.Text.Length >= 3)  // *************** 14/12/2024)
                    {
                        string[] vuelto = oFEgres.ValiCtaCon(Tx_ctaDes.Text, "OMG", "algo");
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
            if ((Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION") && chk_datSimil.Checked == false)
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
        private void Tx_asignado_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_asignado.Text.Trim() != "" && Tx_asignado.Text.Length >= 3)  // *************** 14/12/2024
                {
                    DataRow[] row = Program.dt_definic.Select("idtabella='CMN' and descrizione='" + Tx_asignado.Text.Trim() + "'");
                    foreach (DataRow dat in row)    // .Select("idtabella='CMN' and descrizionerid='" + Tx_asignado.Text.Trim() + "'")
                    {
                        Oasig.codigo = dat[1].ToString();   // codigo
                        Oasig.nombre = dat[2].ToString();    // corto  dat[3].ToString()
                    }
                    if (Oasig.codigo == null || Oasig.codigo == "")
                    {
                        Oasig.nombre = "";
                        Tx_asignado.Text = "";
                        MessageBox.Show("No existe el nombre asignado","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void Tx_ctaDes_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                if (Tx_ctaDes.Text.Trim() != "" && Tx_ctaDes.Text.Length >= 3)  // *************** 14/12/2024
                {
                    string[] vuelto = oFEgres.ValiCtaCon(Tx_ctaDes.Text, "OMG", "algo");
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
        private void Tx_combus_Leave(object sender, EventArgs e)
        {
            decimal monti = 0;
            decimal.TryParse(Tx_combus.Text, out monti);
            Tx_combus.Text = Math.Round(monti, 2).ToString("#,##0.00");
            sumador();
        }
        private void Tx_viati_Leave(object sender, EventArgs e)
        {
            decimal monti = 0;
            decimal.TryParse(Tx_viati.Text, out monti);
            Tx_viati.Text = Math.Round(monti, 2).ToString("#,##0.00");
            sumador();
        }
        private void Tx_rptos_Leave(object sender, EventArgs e)
        {
            decimal monti = 0;
            decimal.TryParse(Tx_rptos.Text, out monti);
            Tx_rptos.Text = Math.Round(monti, 2).ToString("#,##0.00");
            if (chk_dol.CheckState == CheckState.Checked) 
            { 
                tx_dat_rptosS.Text = Math.Round(monti*decimal.Parse(tx_tipcam.Text), 2).ToString("#,##0.00"); 
            }
            else { tx_dat_rptosS.Text = Tx_rptos.Text; }
            sumador();
        }
        private void Tx_varios_Leave(object sender, EventArgs e)
        {
            decimal monti = 0;
            decimal.TryParse(Tx_varios.Text, out monti);
            Tx_varios.Text = Math.Round(monti, 2).ToString("#,##0.00");
            sumador();
        }
        private void Tx_impues_Leave(object sender, EventArgs e)
        {
            decimal impue = 0;
            decimal.TryParse(Tx_impues.Text, out impue);
            Tx_impues.Text = Math.Round(impue, 2).ToString("#,##0.00");
            sumador();
        }
        private void Tx_honor_Leave(object sender, EventArgs e)
        {
            decimal monti = 0;
            decimal.TryParse(Tx_honor.Text, out monti);
            Tx_honor.Text = Math.Round(monti, 2).ToString("#,##0.00");
            sumador();
        }
        private void chk_dol_CheckStateChanged(object sender, EventArgs e)
        {
            var colno = Tx_varios.ForeColor;
            var fonAct = Tx_rptos.Font;
            if (chk_dol.CheckState == CheckState.Checked)
            {
                if (decimal.Parse(tx_tipcam.Text) > 0)
                {
                    decimal monti = 0;
                    decimal.TryParse(Tx_rptos.Text, out monti);
                    tx_dat_rptosS.Text = Math.Round(monti * decimal.Parse(tx_tipcam.Text), 2).ToString("#,##0.00");
                    sumador();
                    Tx_rptos.ForeColor = Color.DarkOliveGreen;
                    Tx_rptos.Font = new Font("Arial", 10, FontStyle.Bold);
                }
                else
                {
                    MessageBox.Show("Debe registrar el tipo de cambio","Atención",MessageBoxButtons.OK,MessageBoxIcon.Hand);
                    chk_dol.CheckState = CheckState.Unchecked;
                    Tx_rptos.Font = new Font("Verdana", 8, FontStyle.Regular); // este es el font normal
                    Tx_rptos.ForeColor = colno;                                // con el color normal
                    tx_tipcam.Focus();
                    return;
                }
            }
            else 
            { 
                tx_dat_rptosS.Text = Tx_rptos.Text;
                sumador();
                Tx_rptos.Font = new Font("Verdana", 8, FontStyle.Regular); // este es el font normal
                Tx_rptos.ForeColor = colno;                                // con el color normal
            }
        }
        private void tx_tipcam_Leave(object sender, EventArgs e)
        {
            Tx_rptos_Leave(null, null);
        }
        private void Tx_asignado_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
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
            limpiaObj();
            limpiaTE();
            escribe("");
            jalaGrilla(diasAtroya, "");
            selecFecha1.Enabled = true;
            Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
            tx_anno.Text = DateTime.Now.Date.Year.ToString();
            tx_anno.ReadOnly = true;
            tx_idOper.ReadOnly = true;
            if (tx_tipcam.Text == "") tx_tipcam.Focus();
            else Tx_asignado.Focus();
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.save_negro40;
            Tx_modo.Text = "EDICION";
            limpiaObj();
            limpiaTE();
            escribe("EDICION");    // sololee("")
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            jalaGrilla(diasAtroya, "");
            tx_idOper.Focus();
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = Conti3.Properties.Resources.delete40;
            Tx_modo.Text = "BORRAR";
            limpiaObj();
            limpiaTE();
            sololee("");
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            jalaGrilla(diasAtroya, "");
            tx_idOper.Focus();
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            Bt_graba.Image = null;
            Tx_modo.Text = "VISUALIZAR";
            limpiaObj();
            limpiaTE();
            sololee("");
            tx_anno.Text = DateTime.Now.Year.ToString();
            tx_anno.ReadOnly = false;
            tx_idOper.ReadOnly = false;
            jalaGrilla(diasAtroya, "");
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

        #region grabaciones
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
            #region validaciones
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDICION")
            {
                
                if (Tx_asignado.Text.Trim() == "")
                {
                    MessageBox.Show("Ingrese el nombre de la persona asignada", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_asignado.Focus();
                    return;
                }
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
                if (tx_montoS.Text == "")
                {
                    MessageBox.Show("Debe ingresar valores de gasto", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_combus.Focus();
                    return;
                }
                else
                {
                    if (decimal.Parse(tx_montoS.Text) <= 0)
                    {
                        MessageBox.Show("Debe ingresar valores de gasto","Atención",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                        Tx_combus.Focus();
                        return;
                    }
                }
                // *************** 14/12/2024
                if (Tx_asignado.Text.Trim() != Oasig.nombre)
                {
                    MessageBox.Show("El asignado no esta completo", "Error, Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_asignado.Select();
                    return;
                }
                if (Tx_ctaDes.Text.Trim() != Ocajd.largo)
                {
                    MessageBox.Show("Cuenta de destino incompleta!", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Tx_ctaDes.Focus();
                    return;
                }
            }
            #endregion
            if (Tx_modo.Text == "NUEVO")
            {
                var aa = MessageBox.Show("Desea grabar el registro actual?","Confirme por favor",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                if (aa == DialogResult.Yes)
                {
                    graba();
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
                    limpiaObj();
                    limpiaTE();
                }
                else { return; }
            }
            if (Tx_modo.Text == "BORRAR")
            {
                // validamos que exista registro que borrar
                if (tx_idOper.Text == "")
                {
                    MessageBox.Show("No hay registro que borrar!", "Identificador en blanco", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }
                var aaa = MessageBox.Show("Confirma que desea BORRAR el registro?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (aaa == DialogResult.Yes)
                {
                    string tabla = "camion";
                    oFEgres.graba_borrar(tabla, tx_anno.Text, "000000000" + oFEgres.CDerecha(tx_idOper.Text, 6), dt_grilla);
                    limpiaObj();
                    limpiaTE();
                }
            }
            var colAct = Tx_combus.ForeColor;
            var fonAct = Tx_combus.Font;
            Tx_rptos.ForeColor = colAct;
            Tx_rptos.Font = fonAct;
            //
            Tx_asignado.Focus();
        }
        private void graba()
        {
            string corre = "";
            string fecOp = Tx_fecha.Text;
            decimal tipCam = 0;
            decimal valCom = 0;
            decimal valVia = 0;
            decimal valRep = 0;
            decimal valImp = 0;
            decimal valHon = 0;
            decimal valVar = 0;
            decimal.TryParse(tx_tipcam.Text, out tipCam);
            decimal.TryParse(Tx_combus.Text, out valCom);
            decimal.TryParse(Tx_viati.Text, out valVia);
            decimal.TryParse(tx_dat_rptosS.Text, out valRep);   // valor en soles de repuestos
            decimal.TryParse(Tx_varios.Text, out valVar);
            decimal.TryParse(Tx_honor.Text, out valHon);
            decimal.TryParse(Tx_impues.Text, out valImp);
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
                    if (true)
                    {
                        corre = oFEgres.correlativo(conn, "MCM", int.Parse(tx_anno.Text));
                        if (corre != "error" && corre != "")
                        {
                            try
                            {
                                Ocamion.creaCamion(Oasig, corre, fecOp, Ocajd, "", "" ,
                                    Omone, decimal.Parse(tx_tipcam.Text), tx_descrip.Text,
                                    valCom, valVia, valRep, valImp, valHon, valVar, Omonto.monSoles, 
                                    Omonto.monDolar, (chk_dol.CheckState == CheckState.Checked) ? 1 : 0, tx_anno.Text);
                                Ocamion.grabaCamion(conn);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Error en grabar Egreso");
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
                jala_ultimo(oFEgres.CDerecha("00000" + corre, 6));
                jalaoc();   // pinta los datos en la pantalla
                Ocamion.IdOper = "";
                tx_idOper.Clear();
                chk_datSimil.Checked = true;
            }
            else
            {
                limpiaObj();
                limpiaTE();
                selecFecha1.Value = DateTime.Now.Date;
                Tx_fecha.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
                tx_anno.Text = DateTime.Now.Date.Year.ToString();
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
                            string fecOp = Tx_fecha.Text;
                            decimal tipCam = 0;
                            decimal valCom = 0;
                            decimal valVia = 0;
                            decimal valRep = 0;
                            decimal valImp = 0;
                            decimal valHon = 0;
                            decimal valVar = 0;
                            decimal.TryParse(tx_tipcam.Text, out tipCam);
                            decimal.TryParse(Tx_combus.Text, out valCom);
                            decimal.TryParse(Tx_viati.Text, out valVia);
                            decimal.TryParse(tx_dat_rptosS.Text, out valRep);   // valor en soles de repuestos
                            decimal.TryParse(Tx_impues.Text, out valImp);
                            decimal.TryParse(Tx_honor.Text, out valHon);
                            decimal.TryParse(Tx_varios.Text, out valVar);
                            Ocamion.creaCamion(Oasig, "000000000" + oFEgres.CDerecha(tx_idOper.Text, 6), fecOp, Ocajd, "", "",
                                    Omone, decimal.Parse(tx_tipcam.Text), tx_descrip.Text,
                                    valCom, valVia, valRep, valImp, valHon, valVar, Omonto.monSoles,
                                    Omonto.monDolar, (chk_dol.CheckState == CheckState.Checked) ? 1 : 0, tx_anno.Text);
                            Ocamion.EditaCamion(conn, tx_anno.Text, "000000000" + oFEgres.CDerecha(tx_idOper.Text, 6));
                            actFilaEnDataI(dt_grilla, "LIM", tx_idOper.Text);
                        }
                    }
                }
            }
        }
        #endregion

        private void panelGeneral1_Paint(object sender, PaintEventArgs e)
        {

        }

    }
}
