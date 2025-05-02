using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;

namespace Conti3
{
    public class preliminares
    {
        private string tipMovPrin;      // Tipo de movimiento, OMG o PERSONAL
        private int idOper;             // Id de la operación, solo para operaciones <> nuevo
        private string fechOper;        // fecha de la operación
        private decimal tipCamb;        // tipo de cambio de la operación si fue en moneda <> a sol
        private string descrip;         // descripcion de la operacion
        private catEgresos catEgreso;   // Categoría del egreso
        private monedas moneda;         // Moneda de la operación
        private cajDestino cajaDes;     // caja destino de la operación, caja desde donde sale el dinero
        private provees proveedor;      // provedor del servicio o bien, es opcional
        private string idMovim;         // idmovimiento
        private montos monto;           // Valor de la operación
        private giroConto giroC;        // giroConto
        private string operador;        // nombre de usuario que digito el registro preliminar
        private string aprobador;       // usuario con provilegios que aprueba
        private int pagado;             // 0=no pagado, 1=pagado o transferido

        libreria lib = new libreria();
        public preliminares() 
        {

        }
        public string TipMovPrin { get => tipMovPrin; set => tipMovPrin = value; }
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
        public string Operador { get => operador; set => operador = value; }
        public string Aprobador { get => aprobador; set => aprobador = value; }
        public int Pagado { get => pagado; set => pagado = value; }

        public void creaPrelim(string _tipMovPrin, string _fechOper, catEgresos _catEgreso, monedas _moneda, montos _monto,
             decimal _tipCamb, cajDestino _cajaDes, provees _proveedor, string _descrip, string _IdMovim, giroConto _giro,
             string _operad, string _aprobador, int _pagado)
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
            operador = _operad;
            aprobador = _aprobador;
            pagado = _pagado;
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
            operador = "";
            aprobador = "";
            pagado = 0;
        }

        public void grabaPrelim(MySqlConnection conn)
        {
            string tabla = "";
            string consulta = "";
            {
                tabla = "cassaprelim";
                consulta = "insert into " + tabla + " (IDBanco,Anno,IDMovimento,DataMovimento,IDConto,IDCategoria,ImportoDU,ImportoSU," +
                                    "Cambio,Descrizione,IDGiroConto,monori,ctaori,ctades,idanagrafica,tipodesgiro,CodGiro," +
                                    "valorOrig,codimon,nombmon,tcMonOri,digitador,tipoE,pagado," +
                                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                                    "@IDB,@Ann,@IDM,@DMo,@IDCo,@IDCa,@IDU,@ISU," +
                                    "@Cam,@Des,@IDG,@mon,@ctao,@ctad,@idan,@tidgiro,@codGiro," +
                                    "@vOrig,@cmon,@nmon,@tcMO,@digit,@tipe,@paga," +
                                    "@veap,@asd,now(),@dipl,@dipw,@nbna)";
            }
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                micon.Parameters.AddWithValue("@IDB", "LIM");   // este campo viene de donde ??? arreglar
                micon.Parameters.AddWithValue("@Ann", DateTime.Now.Year);   // tabla "contatori" debe autoiniciarse al cambiar el año, poner disparador en el login
                micon.Parameters.AddWithValue("@IDM", IdMovim); // este dato viene del metodo "grabar", '00'+contador de la tabla contatori
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", cajaDes.codigo);
                micon.Parameters.AddWithValue("@IDCa", catEgreso.codigo);
                micon.Parameters.AddWithValue("@IDU", monto.monDolar);      // importe en dolares salida 
                micon.Parameters.AddWithValue("@ISU", monto.monSoles);      // importe en soles salida
                micon.Parameters.AddWithValue("@Cam", monto.tipCOri);       // tipCamb
                micon.Parameters.AddWithValue("@Des", descrip);
                micon.Parameters.AddWithValue("@mon", moneda.siglas);    // codigo de la moneda origen de la operación
                micon.Parameters.AddWithValue("@ctao", ""); // ????
                micon.Parameters.AddWithValue("@ctad", ""); // ????
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
                micon.Parameters.AddWithValue("@tcMO", monto.tipCOri);
                micon.Parameters.AddWithValue("@digit", Operador);
                micon.Parameters.AddWithValue("@tipe", (tipMovPrin == "omg") ? "OMG" : "PER");
                //micon.Parameters.AddWithValue("@aprob", ); cuando se grea por primera vez, no hay aprobador
                micon.Parameters.AddWithValue("@paga", pagado);
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                micon.Parameters.AddWithValue("@asd", Program.vg_user);
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }
        }

        public void EditaPrelim(MySqlConnection conn, string year, string corre)
        {
            string consulta = "";
            string tabla = "cassaprelim";
            {
                consulta = "update " + tabla + " set IDBanco=@IDB,Anno=@Ann,DataMovimento=@DMo,IDConto=@IDCo,IDCategoria=@IDCa," +
                    "ImportoDU=@IDU,ImportoSU=@ISU,Cambio=@Cam,Descrizione=@Des,IDGiroConto=@IDG,monori=@mon,ctaori=@ctao,ctades=@ctad," +
                    "digitador=@digit,aprobador=@aprob,idanagrafica=@idan,tipodesgiro=@tidgiro,CodGiro=@codGiro," +
                    "valorOrig=@vOrig,codimon=@cmon,nombmon=@nmon,tcMonOri=@tcMO,tipoE=@tipe,pagado=@paga," +
                    "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                    "where anno=@year and idmovimento=@corre";  // tipoE=@tipe and 
            }
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                micon.Parameters.AddWithValue("@tipe", (tipMovPrin == "omg") ? "OMG" : "PER");
                micon.Parameters.AddWithValue("@IDB", "LIM");   // este campo viene de donde ??? arreglar
                micon.Parameters.AddWithValue("@Ann", year);   // tabla "contatori" debe autoiniciarse al cambiar el año, poner disparador en el login
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", cajaDes.codigo);
                micon.Parameters.AddWithValue("@IDCa", catEgreso.codigo);
                micon.Parameters.AddWithValue("@IDU", monto.monDolar);      // importe en dolares salida 
                micon.Parameters.AddWithValue("@ISU", monto.monSoles);      // importe en soles salida
                micon.Parameters.AddWithValue("@Cam", monto.tipCOri);       // tipCamb
                micon.Parameters.AddWithValue("@Des", descrip);
                micon.Parameters.AddWithValue("@IDG", giroC.idcod);         //       idgiroconto
                micon.Parameters.AddWithValue("@mon", moneda.siglas);       // codigo de la moneda origen de la operación
                micon.Parameters.AddWithValue("@ctao", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@ctad", ""); // esto va con el giroconto creo
                micon.Parameters.AddWithValue("@digit", Operador);
                micon.Parameters.AddWithValue("@aprob", Aprobador);
                micon.Parameters.AddWithValue("@idan", (proveedor.codigo == null) ? "" : proveedor.codigo);
                micon.Parameters.AddWithValue("@tidgiro", giroC.tipodes);   //      OMG o PER
                micon.Parameters.AddWithValue("@codGiro", (giroC.codigo == null) ? "" : giroC.codigo);    // CodGiro
                micon.Parameters.AddWithValue("@vOrig", monto.monOrige);
                micon.Parameters.AddWithValue("@cmon", moneda.codigo);
                micon.Parameters.AddWithValue("@nmon", moneda.nombre);
                micon.Parameters.AddWithValue("@tcMO", monto.tipCOri);
                micon.Parameters.AddWithValue("@year", year);
                micon.Parameters.AddWithValue("@corre", corre);
                micon.Parameters.AddWithValue("@paga", pagado);
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                micon.Parameters.AddWithValue("@asd", Program.vg_user);
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }
        }

        public void actuaPrelim(MySqlConnection conn, string year, string corre)
        {
            string actua = "update cassaprelim set aprobador=@apro,fecproc=now(),pagado=@paga," +
                "verApp=@veap,diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                "where anno=@Ann and idmovimento=@idm";
            using (MySqlCommand micon = new MySqlCommand(actua, conn))
            {
                micon.Parameters.AddWithValue("@Ann", year);
                micon.Parameters.AddWithValue("@idm", corre);
                micon.Parameters.AddWithValue("@apro", Aprobador);
                micon.Parameters.AddWithValue("@paga", pagado);
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                //micon.Parameters.AddWithValue("@asd", Program.vg_user); -> no se actualiza porque corresponde al posible usuario que edito
                // campo fechm -> no se actualiza porque corresponde a la posible fecha de edicion del registro preliminar
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }
        }
    }
}
