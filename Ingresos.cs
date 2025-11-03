using MySql.Data.MySqlClient;
using System;
using System.Security.Cryptography;
using System.Web.ModelBinding;
using System.Windows.Forms;

namespace Conti3
{
    class Ingresos
    {
        private string tipMovPrin;  // Tipo de movimiento, OMG o PERSONAL
        private int idOper;         // Id de la operación, solo para operaciones <> nuevo
        private string annoOp;      // año de la operacion
        private string fechOper;    // fecha de la operación
        private decimal tipCamb;    // tipo de cambio de la operación si fue en moneda <> a sol
        private string descrip;     // descripcion de la operacion
        private catIngresos catIngreso;   // Categoría del Ingreso
        private monedas moneda;         // Moneda de la operación
        private cajDestino cajaDes;     // caja destino de la operación, caja desde donde sale el dinero
        private string idMovim;         // idmovimiento
        private montos monto;           // Valor de la operación
        private giroConto giroC;        // giroConto
        libreria lib = new libreria();
        public Ingresos()
        {

        }

        public string TipMovPrin { get => tipMovPrin; set => tipMovPrin = value; }
        public int IdOper { get => idOper; set => idOper = value; }
        public string AnnoOp { get => annoOp; set => annoOp = value; }
        public string FechOper { get => fechOper; set => fechOper = value; }
        public monedas Moneda { get => moneda; set => moneda = value; }
        public montos Monto { get => monto; set => monto = value; }
        public decimal TipCamb { get => tipCamb; set => tipCamb = value; }
        public cajDestino CajaDes { get => cajaDes; set => cajaDes = value; }
        public string Descrip { get => descrip; set => descrip = value; }
        public catIngresos CatIngreso { get => catIngreso; set => catIngreso = value; }
        public string IdMovim { get => idMovim; set => idMovim = value; }
        public giroConto GiroC { get => giroC; set => giroC = value; }

        public void creaIngreso(string _tipMovPrin, string _fechOper, catIngresos _catIngreso, monedas _moneda, montos _monto, decimal _tipCamb,
            cajDestino _cajaDes, string _descrip, string _IdMovim, giroConto _giro, string _anop)
        {
            tipMovPrin = _tipMovPrin;
            annoOp = _anop;
            fechOper = _fechOper;
            catIngreso = _catIngreso;
            moneda = _moneda;
            monto = _monto;
            tipCamb = _tipCamb;
            cajaDes = _cajaDes;
            descrip = _descrip;
            idMovim = _IdMovim;
            giroC = _giro;
        }

        public void limpia()
        {
            tipMovPrin = "";
            fechOper = "";
            catIngreso = null;
            moneda = null;
            monto = null;
            tipCamb = 0;
            cajaDes = null;
            descrip = "";
            idMovim = "";
            giroC = null;
            annoOp = "";
        }

        public void grabaIngreso(MySqlConnection conn)
        {
            string tabla = "";
            string consulta = "";
            if (tipMovPrin == "omg")
            {
                tabla = "cassaomg";
                consulta = "insert into " + tabla + " (IDBanco,Anno,IDMovimento,DataMovimento,IDDestino,IDCategoria,ImportoDE,ImportoSE," +
                                    "Cambio,Descrizione,IDGiroConto,monori,ctaori,ctades,usuario,dia,tipodesgiro,CodGiro," +
                                    "valorOrig,codimon,nombmon,tcMonOri," +
                                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                                    "@IDB,@Ann,@IDM,@DMo,@IDCo,@IDCa,@IDE,@ISE,@Cam,@Des,@IDG,@mon,@ctao,@ctad,@usua,now(),@tidgiro,@codGiro," +
                                    "@vOrig,@cmon,@nmon,@tcMO," +
                                    "@veap,@asd,now(),@dipl,@dipw,@nbna)";
            }
            else
            {
                tabla = "cassaconti";
                consulta = "insert into " + tabla + " (IDBanco,Anno,IDMovimento,DataMovimento,IDConto,IDCategoria,ImportoDE,ImportoSE," +
                                    "Cambio,Descrizione,IDGiroConto,monori,ctaori,ctades,usuario,dia,tipodesgiro,CodGiro," +
                                    "valorOrig,codimon,nombmon,tcMonOri," +
                                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                                    "@IDB,@Ann,@IDM,@DMo,@IDCo,@IDCa,@IDE,@ISE,@Cam,@Des,@IDG,@mon,@ctao,@ctad,@usua,now(),@tidgiro,@codGiro," +
                                    "@vOrig,@cmon,@nmon,@tcMO," +
                                    "@veap,@asd,now(),@dipl,@dipw,@nbna)";
            }
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                //micon.Parameters.AddWithValue("@tab", tabla);
                micon.Parameters.AddWithValue("@IDB", "LIM");   // este campo viene de donde ??? arreglar
                micon.Parameters.AddWithValue("@Ann", AnnoOp);   // DateTime.Now.Year ...  tabla "contatori" debe autoiniciarse al cambiar el año, poner disparador en el login
                micon.Parameters.AddWithValue("@IDM", idMovim); 
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", cajaDes.codigo);
                micon.Parameters.AddWithValue("@IDCa", catIngreso.codigo);
                micon.Parameters.AddWithValue("@IDE", monto.monDolar);    // importe en dolares entrada
                micon.Parameters.AddWithValue("@ISE", monto.monSoles);    // importe en soles entrada
                //micon.Parameters.AddWithValue("@IEE", 0);                // importe en euros entrada
                //micon.Parameters.AddWithValue("@IEU", monto.monEuros);   // importe en euros salida 
                micon.Parameters.AddWithValue("@Cam", TipCamb);     // monto.tipCOri
                micon.Parameters.AddWithValue("@Des", descrip);
                micon.Parameters.AddWithValue("@mon", Moneda.siglas);     // codigo de la moneda origen de la operación  monto.codMOrige
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
                    micon.Parameters.AddWithValue("@IDG", giroC.idcod);        // cuenta destino del giroconto
                    micon.Parameters.AddWithValue("@tidgiro", giroC.tipodes);   // tipo de cuenta destino del giro, OMG o PERSONAL
                    micon.Parameters.AddWithValue("@codGiro", (giroC.codigo == null) ? "" : giroC.codigo);    // CodGiro
                }
                micon.Parameters.AddWithValue("@vOrig", monto.monOrige);
                micon.Parameters.AddWithValue("@cmon", moneda.codigo);
                micon.Parameters.AddWithValue("@nmon", moneda.nombre);
                micon.Parameters.AddWithValue("@tcMO", TipCamb);  // monto.tipCOri
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                micon.Parameters.AddWithValue("@asd", Program.vg_user);
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }
        }
        public void EditaIngreso(MySqlConnection conn, string year, string corre)
        {
            string consulta = "";
            if (tipMovPrin == "omg")
            {
                consulta = "update cassaomg set IDBanco=@IDB,Anno=@Ann,DataMovimento=@DMo,IDDestino=@IDCo,IDCategoria=@IDCa," +
                    "ImportoDE=@IDU,ImportoSE=@ISU,Cambio=@Cam,Descrizione=@Des,IDGiroConto=@IDG,monori=@mon,ctaori=@ctao,ctades=@ctad," +
                    "usuario=@usua,diaM=now(),tipodesgiro=@tidgiro,valorOrig=@vOrig,codimon=@cmon,nombmon=@nmon,tcMonOri=@tcMO," +
                    "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                    "where anno=@year and idmovimento=@corre";
            }
            else
            {
                consulta = "update cassaconti set IDBanco=@IDB,Anno=@Ann,DataMovimento=@DMo,IDConto=@IDCo,IDCategoria=@IDCa," +
                    "ImportoDE=@IDU,ImportoSE=@ISU,Cambio=@Cam,Descrizione=@Des,IDGiroConto=@IDG,monori=@mon,ctaori=@ctao,ctades=@ctad," +
                    "usuario=@usua,diaM=now(),tipodesgiro=@tidgiro,valorOrig=@vOrig,codimon=@cmon,nombmon=@nmon,tcMonOri=@tcMO," +
                    "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                    "where anno=@year and idmovimento=@corre";
            }
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                micon.Parameters.AddWithValue("@IDB", "LIM");   // este campo viene de donde ??? arreglar
                micon.Parameters.AddWithValue("@Ann", year);    // DateTime.Now.Year .. tabla "contatori" debe autoiniciarse al cambiar el año, poner disparador en el login
                //micon.Parameters.AddWithValue("@IDM", IdMovim); // este dato viene del metodo "grabar", '00'+contador de la tabla contatori
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", cajaDes.codigo);
                micon.Parameters.AddWithValue("@IDCa", catIngreso.codigo);
                //micon.Parameters.AddWithValue("@IEE", 0);                   // importe en euros entrada
                micon.Parameters.AddWithValue("@IDU", monto.monDolar);    // importe en dolares salida 
                micon.Parameters.AddWithValue("@ISU", monto.monSoles);    // importe en soles salida
                //micon.Parameters.AddWithValue("@IEU", monto.monEuros);    // importe en euros salida 
                micon.Parameters.AddWithValue("@Cam", TipCamb);             // tipo de cambio  monto.tipCOri
                micon.Parameters.AddWithValue("@Des", descrip);
                micon.Parameters.AddWithValue("@IDG", GiroC.idcod);  // cuenta origen del GIRO
                micon.Parameters.AddWithValue("@mon", moneda.siglas);    // codigo de la moneda origen de la operación
                micon.Parameters.AddWithValue("@ctao", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@ctad", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@usua", Program.vg_user);
                micon.Parameters.AddWithValue("@tidgiro", GiroC.tipodes);   // tipo de giroconto OMG | PER
                micon.Parameters.AddWithValue("@vOrig", monto.monOrige);
                micon.Parameters.AddWithValue("@cmon", moneda.codigo);
                micon.Parameters.AddWithValue("@nmon", moneda.nombre);
                micon.Parameters.AddWithValue("@tcMO", TipCamb);   // monto.tipCOri
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
