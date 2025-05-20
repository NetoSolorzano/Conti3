using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Conti3
{
    public class Egresos
    {
        private string tipMovPrin;  // Tipo de movimiento, OMG o PERSONAL
        private string annoOp;      // año
        private int idOper;         // Id de la operación, solo para operaciones <> nuevo
        private string fechOper;    // fecha de la operación
        private decimal tipCamb;    // tipo de cambio de la operación si fue en moneda <> a sol
        private string descrip;     // descripcion de la operacion
        private catEgresos catEgreso;   // Categoría del egreso
        private monedas moneda;         // Moneda de la operación
        private cajDestino cajaDes;     // caja destino de la operación, caja desde donde sale el dinero
        private provees proveedor;      // provedor del servicio o bien, es opcional
        private string idMovim;         // idmovimiento
        private montos monto;           // Valor de la operación
        private giroConto giroC;        // giroConto
        libreria lib = new libreria();  
        public Egresos()
        {

        }

        public string TipMovPrin { get => tipMovPrin; set => tipMovPrin = value; }
        public string AnnoOp { get => annoOp; set => annoOp = value; }
        public int IdOper { get => idOper; set => idOper = value; }
        public string FechOper { get => fechOper; set => fechOper = value; }
        public monedas Moneda { get => moneda; set => moneda = value; }
        public montos Monto { get => monto; set => monto = value; }
        public decimal TipCamb { get => tipCamb; set => tipCamb = value; }
        public cajDestino CajaDes { get => cajaDes; set => cajaDes = value; }
        public provees Proveedor { get => proveedor; set => proveedor = value; }
        public string Descrip { get => descrip; set => descrip = value; }
        public catEgresos CatEgreso { get => catEgreso; set => catEgreso = value; }
        public string IdMovim { get => idMovim; set => idMovim = value; }
        public giroConto GiroC { get => giroC; set => giroC = value; }

        public void creaEgreso(string _tipMovPrin, string _fechOper, catEgresos _catEgreso, monedas _moneda, montos _monto, decimal _tipCamb, 
            cajDestino _cajaDes, provees _proveedor, string _descrip, string _IdMovim, giroConto _giro, string _yea)   // , string _IdMovim <- este dato se genera dentro del trigger beforeinsert
        {
            tipMovPrin = _tipMovPrin;
            fechOper = _fechOper;
            catEgreso = _catEgreso;
            moneda = _moneda;
            monto = _monto;
            tipCamb = _tipCamb;
            cajaDes = _cajaDes;
            proveedor = _proveedor;
            descrip = _descrip;
            idMovim = _IdMovim;
            giroC = _giro;
            annoOp = _yea;
        }
        public void limpia()
        {
            tipMovPrin = "";
            fechOper = "";
            catEgreso = null;
            moneda = null;
            monto = null;
            tipCamb = 0;
            cajaDes = null;
            proveedor = null;
            descrip = "";
            idMovim = "";
            giroC = null;
            annoOp = "";
        }
        public void grabaEgreso(MySqlConnection conn)
        {
            string tabla = "";
            string consulta = "";
            if (tipMovPrin == "omg")
            {
                tabla = "cassaomg";
                consulta = "insert into " + tabla + " (IDBanco,Anno,IDMovimento,DataMovimento,IDDestino,IDCategoria,ImportoDU,ImportoSU," +
                                    "Cambio,Descrizione,IDGiroConto,monori,ctaori,ctades,usuario,dia,idanagrafica,tipodesgiro,CodGiro," +
                                    "valorOrig,codimon,nombmon,tcMonOri," +
                                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                                    "@IDB,@Ann,@IDM,@DMo,@IDCo,@IDCa,@IDU,@ISU,@Cam,@Des,@IDG,@mon,@ctao,@ctad,@usua,now(),@idan,@tidgiro,@codGiro," +
                                    "@vOrig,@cmon,@nmon,@tcMO," +
                                    "@veap,@asd,now(),@dipl,@dipw,@nbna)";
            }
            else
            {
                tabla = "cassaconti";
                consulta = "insert into " + tabla + " (IDBanco,Anno,IDMovimento,DataMovimento,IDConto,IDCategoria,ImportoDU,ImportoSU," +
                                    "Cambio,Descrizione,IDGiroConto,monori,ctaori,ctades,usuario,dia,idanagrafica,tipodesgiro,CodGiro," +
                                    "valorOrig,codimon,nombmon,tcMonOri," +
                                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                                    "@IDB,@Ann,@IDM,@DMo,@IDCo,@IDCa,@IDU,@ISU,@Cam,@Des,@IDG,@mon,@ctao,@ctad,@usua,now(),@idan,@tidgiro,@codGiro," +
                                    "@vOrig,@cmon,@nmon,@tcMO," +
                                    "@veap,@asd,now(),@dipl,@dipw,@nbna)";
            }
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                ///micon.Parameters.AddWithValue("@tab", tabla);
                micon.Parameters.AddWithValue("@IDB", "LIM");   // este campo viene de donde ??? arreglar
                micon.Parameters.AddWithValue("@Ann", AnnoOp);   // DateTime.Now.Year  ...   tabla "contatori" debe autoiniciarse al cambiar el año, poner disparador en el login
                micon.Parameters.AddWithValue("@IDM", IdMovim); // este dato viene del metodo "grabar", '00'+contador de la tabla contatori
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", cajaDes.codigo);
                micon.Parameters.AddWithValue("@IDCa", catEgreso.codigo);
                //micon.Parameters.AddWithValue("@IEE", 0);                 // importe en euros entrada
                micon.Parameters.AddWithValue("@IDU", monto.monDolar);      // importe en dolares salida 
                micon.Parameters.AddWithValue("@ISU", monto.monSoles);      // importe en soles salida
                //micon.Parameters.AddWithValue("@IEU", monto.monEuros);    // importe en euros salida 
                micon.Parameters.AddWithValue("@Cam", TipCamb);       // tipCamb   monto.tipCOri
                micon.Parameters.AddWithValue("@Des", descrip);
                micon.Parameters.AddWithValue("@mon", moneda.siglas);    // codigo de la moneda origen de la operación
                micon.Parameters.AddWithValue("@ctao", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@ctad", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@usua", Program.vg_user);
                if (giroC.ctades == null || giroC.ctades == "")
                {
                    micon.Parameters.AddWithValue("@IDG", "");          // cuenta destino del giroconto
                    micon.Parameters.AddWithValue("@tidgiro", "");      // tipo de cuenta destino del giro, OMG o PERSONAL
                    micon.Parameters.AddWithValue("@codGiro", "");
                }
                else
                {
                    micon.Parameters.AddWithValue("@IDG", giroC.idcod);         // giroC.ctades   cuenta destino del giroconto
                    micon.Parameters.AddWithValue("@tidgiro", giroC.tipodes);   // tipo de cuenta destino del giro, OMG o PERSONAL
                    micon.Parameters.AddWithValue("@codGiro", (giroC.codigo == null) ? "" : giroC.codigo);    // CodGiro
                }
                micon.Parameters.AddWithValue("@idan", (proveedor.codigo == null) ? "" : proveedor.codigo);
                micon.Parameters.AddWithValue("@vOrig", monto.monOrige);
                micon.Parameters.AddWithValue("@cmon", moneda.codigo);
                micon.Parameters.AddWithValue("@nmon", moneda.nombre);
                micon.Parameters.AddWithValue("@tcMO", TipCamb);   // monto.tipCOri
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                micon.Parameters.AddWithValue("@asd", Program.vg_user);
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }
        }
        public void EditaEgreso(MySqlConnection conn, string year, string corre)
        {
            string consulta = "";
            if (tipMovPrin == "omg")
            {
                consulta = "update cassaomg set IDBanco=@IDB,Anno=@Ann,DataMovimento=@DMo,IDDestino=@IDCo,IDCategoria=@IDCa," +
                    "ImportoDU=@IDU,ImportoSU=@ISU,Cambio=@Cam,Descrizione=@Des,IDGiroConto=@IDG,monori=@mon,ctaori=@ctao,ctades=@ctad," +
                    "usuario=@usua,diaM=now(),idanagrafica=@idan,tipodesgiro=@tidgiro,valorOrig=@vOrig,codimon=@cmon,nombmon=@nmon,tcMonOri=@tcMO," +
                    "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                    "where anno=@year and idmovimento=@corre";
            }
            else
            {
                consulta = "update cassaconti set IDBanco=@IDB,Anno=@Ann,DataMovimento=@DMo,IDConto=@IDCo,IDCategoria=@IDCa," +
                    "ImportoDU=@IDU,ImportoSU=@ISU,Cambio=@Cam,Descrizione=@Des,IDGiroConto=@IDG,monori=@mon,ctaori=@ctao,ctades=@ctad," +
                    "usuario=@usua,diaM=now(),idanagrafica=@idan,tipodesgiro=@tidgiro,valorOrig=@vOrig,codimon=@cmon,nombmon=@nmon,tcMonOri=@tcMO," +
                    "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                    "where anno=@year and idmovimento=@corre";
            }
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                micon.Parameters.AddWithValue("@IDB", "LIM");   // este campo viene de donde ??? arreglar
                micon.Parameters.AddWithValue("@Ann", year);   // tabla "contatori" debe autoiniciarse al cambiar el año, poner disparador en el login
                //micon.Parameters.AddWithValue("@IDM", IdMovim); // este dato viene del metodo "grabar", '00'+contador de la tabla contatori
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", cajaDes.codigo);
                micon.Parameters.AddWithValue("@IDCa", catEgreso.codigo);
                //micon.Parameters.AddWithValue("@IEE", 0);                 // importe en euros entrada
                micon.Parameters.AddWithValue("@IDU", monto.monDolar);      // importe en dolares salida 
                micon.Parameters.AddWithValue("@ISU", monto.monSoles);      // importe en soles salida
                //micon.Parameters.AddWithValue("@IEU", monto.monEuros);    // importe en euros salida 
                micon.Parameters.AddWithValue("@Cam", TipCamb);       // tipCamb   monto.tipCOri
                micon.Parameters.AddWithValue("@Des", descrip);
                micon.Parameters.AddWithValue("@IDG", giroC.idcod);         // idgiroconto
                micon.Parameters.AddWithValue("@mon", moneda.siglas);       // codigo de la moneda origen de la operación
                micon.Parameters.AddWithValue("@ctao", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@ctad", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@usua", Program.vg_user);
                micon.Parameters.AddWithValue("@idan", proveedor.codigo);
                micon.Parameters.AddWithValue("@tidgiro", giroC.tipodes);   // OMG o PER
                micon.Parameters.AddWithValue("@vOrig", monto.monOrige);
                micon.Parameters.AddWithValue("@cmon", moneda.codigo);
                micon.Parameters.AddWithValue("@nmon", moneda.nombre);
                micon.Parameters.AddWithValue("@tcMO", TipCamb);      // monto.tipCOri
                micon.Parameters.AddWithValue("@year", year);
                micon.Parameters.AddWithValue("@corre", corre);
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                micon.Parameters.AddWithValue("@asd", Program.vg_user);
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }

        }
    }
}
