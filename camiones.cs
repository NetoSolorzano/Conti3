using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;

namespace Conti3
{
    public class camiones
    {
        private provees placa;          // idcodice de la tabla desc_cmn
        private string idOper;
        private string annoOp;          // año de la operación
        private string fechOper;        // fecha de la operación
        private cajDestino cajaDes;     // caja destino de la operación, caja desde donde sale el dinero
        private string bcoOrigen;
        private string bcoDestin;
        private monedas moneda;         // Moneda de la operación
        private decimal tipCamb;    // tipo de cambio de la operación si fue en moneda <> a sol
        private string descrip;     // descripcion de la operacion
        private decimal combust;    // importe gasto combustible
        private decimal viaticos;   // importe gasto viativos
        private decimal respuest;   // importe gastos repuestos
        private decimal honorar;    // gastos de honorarios 
        private decimal impuests;   // gastos por impuestos del camion
        private decimal varios;     // gastos varios
        private decimal totalS;     // gran total en soles
        private decimal totalD;     // gran total en dolares
        private int rptsDol;        // repuestos en dolares 1=si
        libreria lib = new libreria();
        public camiones()
        {

        }

        public provees Placa { get => placa; set => placa = value; }
        public string IdOper { get => idOper; set => idOper = value; }
        public string AnnoOp { get => annoOp; set => annoOp = value; }
        public string FechOper { get => fechOper; set => fechOper = value; }
        public cajDestino CajaDes { get => cajaDes; set => cajaDes = value; }
        public string BcoOrigen { get => bcoOrigen; set => bcoOrigen = value; }
        public string BcoDestin { get => bcoDestin; set => bcoDestin = value; }
        public monedas Moneda { get => moneda; set => moneda = value; }
        public decimal TipCamb { get => tipCamb; set => tipCamb = value; }
        public string Descrip { get => descrip; set => descrip = value; }
        public decimal Combust { get => combust; set => combust = value; }
        public decimal Viaticos { get => viaticos; set => viaticos = value; }
        public decimal Respuest { get => respuest; set => respuest = value; }
        public decimal Honorar { get => honorar; set => honorar = value; }
        public decimal Impuests { get => impuests; set => impuests = value; }
        public decimal Varios { get => varios; set => varios = value; }
        public decimal TotalS { get => totalS; set => totalS = value; }
        public decimal TotalD { get => totalD; set => totalD = value; }
        public int RptsDol { get => rptsDol; set => rptsDol = value; }

        public void creaCamion(provees _placa, string _idoper, string _fecha, cajDestino _destino,
            string _bcoOrigen, string _bcoDestin, monedas _moneda, decimal _tipCamb, string _descrip,
            decimal _combust, decimal _viaticos, decimal _respuest, decimal _impuests, decimal _honorar,
            decimal _varios, decimal _totalS, decimal _totalD, int _repDol, string _anOp)
        {
            placa = _placa;
            idOper = _idoper;
            annoOp = _anOp;
            fechOper = _fecha;
            cajaDes = _destino;
            bcoOrigen = _bcoOrigen;
            bcoDestin = _bcoDestin;
            moneda = _moneda;
            tipCamb = _tipCamb;
            descrip = _descrip;
            combust = _combust;
            viaticos = _viaticos;
            respuest = _respuest;
            honorar = _honorar;
            impuests = _impuests;
            varios = _varios;
            totalS = _totalS;
            totalD = _totalD;
            rptsDol = _repDol;
        }
        public void limpia()
        {
            placa = null;
            idOper = "";
            annoOp = "";
            fechOper = "";        // fecha de la operación
            cajaDes = null;     // caja destino de la operación, caja desde donde sale el dinero
            bcoOrigen = "";
            bcoDestin = "";
            moneda = null;         // Moneda de la operación
            tipCamb = 0;    // tipo de cambio de la operación si fue en moneda <> a sol
            descrip = "";   // descripcion de la operacion
            combust = 0;    // gasto combustible
            viaticos = 0;   // gasto viaticos
            respuest = 0;   // gastos repuestos
            honorar = 0;    // gastos de honorarios 
            impuests = 0;   // gastos por impuestos del camion
            varios = 0;     // gastos varios
            totalS = 0;
            totalD = 0;
        }
        public void grabaCamion(MySqlConnection conn)
        {
            decimal combD = combust / tipCamb;
            decimal viatD = viaticos / tipCamb;
            decimal respD = respuest / tipCamb;
            decimal variD = varios / tipCamb;
            decimal honoD = honorar / tipCamb;
            decimal impuD = impuests / tipCamb;
            string consulta = "INSERT INTO camion (IDBanco,Anno,IDMovimento,DataMovimento,IDCamion,IDDestino," +
                "ImpCarbD,ImpViaD,ImpRicD,ImpVariD,imphond,ImpImpD,ImpCarbS,ImpViaS,ImpRicS,ImpVariS,imphons,ImpImpS," +
                "ImpTotD,ImpTotS,Cambio,Descrizione,numero,monori,ctaori," +
                "ctades,usuario,dia,repdol,codimon,nombmon," +
                "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                "@IDB,@Ann,@IDM,@DMo,@IDCo,@IDCa," +
                "@ImpCarbD,@ImpViaD,@ImpRicD,@ImpVariD,@imphond,@ImpImpD,@ImpCarbS,@ImpViaS,@ImpRicS,@ImpVariS,@imphons,@ImpImpS," +
                "@ImpTotD,@ImpTotS,@Cambio,@Descr,@numero,@monori,@ctaori," +
                "@ctades,@usuario,now(),@repdol,@codimon,@nombmon," +
                "@veap,@asd,now(),@dipl,@dipw,@nbna)";
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                ///micon.Parameters.AddWithValue("@tab", tabla);
                micon.Parameters.AddWithValue("@IDB", "LIM"); 
                micon.Parameters.AddWithValue("@Ann", annoOp);    // int.Parse(fechOper.Substring(6, 4))    Tx_rptos.ForeColor = colAct;
                micon.Parameters.AddWithValue("@IDM", idOper);
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", placa.codigo);
                micon.Parameters.AddWithValue("@IDCa", cajaDes.codigo);
                micon.Parameters.AddWithValue("@ImpCarbD", combD);
                micon.Parameters.AddWithValue("@ImpViaD", viatD);
                micon.Parameters.AddWithValue("@ImpRicD", respD);
                micon.Parameters.AddWithValue("@ImpVariD", variD);
                micon.Parameters.AddWithValue("@imphond", honoD);
                micon.Parameters.AddWithValue("@ImpImpD", impuD);
                micon.Parameters.AddWithValue("@ImpCarbS", combust);
                micon.Parameters.AddWithValue("@ImpViaS", viaticos);
                micon.Parameters.AddWithValue("@ImpRicS", respuest);
                micon.Parameters.AddWithValue("@ImpVariS", varios);
                micon.Parameters.AddWithValue("@imphons", honorar);
                micon.Parameters.AddWithValue("@ImpImpS", impuests);
                micon.Parameters.AddWithValue("@Cambio", tipCamb);
                micon.Parameters.AddWithValue("@ImpTotD", totalD);
                micon.Parameters.AddWithValue("@ImpTotS", totalS);
                micon.Parameters.AddWithValue("@Descr", descrip);
                micon.Parameters.AddWithValue("@numero", 1);
                micon.Parameters.AddWithValue("@monori", moneda.siglas);
                micon.Parameters.AddWithValue("@ctaori", "");
                micon.Parameters.AddWithValue("@ctades", "");
                micon.Parameters.AddWithValue("@usuario", Program.vg_user);
                micon.Parameters.AddWithValue("@repdol", rptsDol);
                micon.Parameters.AddWithValue("@codimon", moneda.codigo);
                micon.Parameters.AddWithValue("@nombmon", moneda.nombre);
                micon.Parameters.AddWithValue("@veap", System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion);
                micon.Parameters.AddWithValue("@asd", Program.vg_user);
                micon.Parameters.AddWithValue("@dipl", lib.iplan());
                micon.Parameters.AddWithValue("@dipw", Conti3.Program.vg_ipwan);
                micon.Parameters.AddWithValue("@nbna", Environment.MachineName);
                micon.ExecuteNonQuery();
            }
        }
        public void EditaCamion(MySqlConnection conn, string year, string corre)
        {
            string consulta = "";
            decimal combD = combust / tipCamb;
            decimal viatD = viaticos / tipCamb;
            decimal respD = respuest / tipCamb;
            decimal variD = varios / tipCamb;
            decimal honoD = honorar / tipCamb;
            decimal impuD = impuests / tipCamb;
            consulta = "UPDATE camion SET IDBanco=@IDB,DataMovimento=@DMo,IDCamion=@IDCo,IDDestino=@IDCa," +
                "ImpCarbD=@ImpCarbD,ImpViaD=@ImpViaD,ImpRicD=@ImpRicD,ImpVariD=@ImpVariD,imphond=@imphond," +
                "ImpImpD=@ImpImpD,ImpCarbS=@ImpCarbS,ImpViaS=@ImpViaS,ImpRicS=@ImpRicS,ImpVariS=@ImpVariS," +
                "imphons=@imphons,ImpImpS=@ImpImpS,ImpTotD=@ImpTotD,ImpTotS=@ImpTotS,Cambio=@Cambio," +
                "Descrizione=@Descr,numero=@numero,monori=@monori,ctaori=@ctaori,ctades=@ctades," +
                "usuario=@usuario,repdol=@repdol,codimon=@codimon,nombmon=@nombmon," +
                "verApp=@veap,userm=@asd,fechm=now(),diriplan4=@dipl,diripwan4=@dipw,netbname=@nbna " +
                "where anno=@Ann and idmovimento=@IDM";
            using (MySqlCommand micon = new MySqlCommand(consulta, conn))
            {
                micon.Parameters.AddWithValue("@IDB", "LIM");
                micon.Parameters.AddWithValue("@Ann", int.Parse(year));
                micon.Parameters.AddWithValue("@IDM", idOper);
                micon.Parameters.AddWithValue("@DMo", fechOper.Substring(6, 4) + "-" + fechOper.Substring(3, 2) + "-" + fechOper.Substring(0, 2));
                micon.Parameters.AddWithValue("@IDCo", placa.codigo);
                micon.Parameters.AddWithValue("@IDCa", cajaDes.codigo);
                micon.Parameters.AddWithValue("@ImpCarbD", combD);
                micon.Parameters.AddWithValue("@ImpViaD", viatD);
                micon.Parameters.AddWithValue("@ImpRicD", respD);
                micon.Parameters.AddWithValue("@ImpVariD", variD);
                micon.Parameters.AddWithValue("@imphond", honoD);
                micon.Parameters.AddWithValue("@ImpImpD", impuD);
                micon.Parameters.AddWithValue("@ImpCarbS", combust);
                micon.Parameters.AddWithValue("@ImpViaS", viaticos);
                micon.Parameters.AddWithValue("@ImpRicS", respuest);
                micon.Parameters.AddWithValue("@ImpVariS", varios);
                micon.Parameters.AddWithValue("@imphons", honorar);
                micon.Parameters.AddWithValue("@ImpImpS", impuests);
                micon.Parameters.AddWithValue("@Cambio", tipCamb);
                micon.Parameters.AddWithValue("@ImpTotD", totalD);
                micon.Parameters.AddWithValue("@ImpTotS", totalS);
                micon.Parameters.AddWithValue("@Descr", descrip);
                micon.Parameters.AddWithValue("@numero", 1);
                micon.Parameters.AddWithValue("@monori", moneda.siglas);
                micon.Parameters.AddWithValue("@ctaori", "");
                micon.Parameters.AddWithValue("@ctades", "");
                micon.Parameters.AddWithValue("@usuario", Program.vg_user);
                micon.Parameters.AddWithValue("@repdol", rptsDol);
                micon.Parameters.AddWithValue("@codimon", moneda.codigo);
                micon.Parameters.AddWithValue("@nombmon", moneda.nombre);
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
