using CrystalDecisions.CrystalReports.Engine;
using System.Windows.Forms;

namespace Conti3
{
    public partial class FrmVisual : Form
    {
        DataSet1 _dataSet1;
        public FrmVisual()
        {
            InitializeComponent();
        }
        public FrmVisual(DataSet1 dataSet) : this()
        { 
            _dataSet1 = dataSet;
        }
        private void FrmVisual_Load(object sender, System.EventArgs e)
        {
            // reporte Movimiento por caja - personal 
            if (_dataSet1.repSaldoIni_.Rows.Count > 0)
            {
                string nf = _dataSet1.repSaldoIni_.Rows[0].ItemArray[1].ToString();
                ReportDocument rpt = new ReportDocument();
                rpt.Load(nf);
                rpt.SetDataSource(_dataSet1);
                //crystalReportViewer1.Width = 76; NO HACE CASO
                //crystalReportViewer1.Height = 108; NO HACE CASO
                crystalReportViewer1.ReportSource = rpt;
            }
            if (_dataSet1.CabGasCam.Rows.Count > 0)
            {
                string nf = _dataSet1.CabGasCam.Rows[0].ItemArray[1].ToString();
                ReportDocument rpt = new ReportDocument();
                rpt.Load(nf);
                rpt.SetDataSource(_dataSet1);
                crystalReportViewer1.ReportSource = rpt;
            }
            if (_dataSet1.ResCtasPers.Rows.Count > 0)
            {
                string nf = _dataSet1.ResCtasPers.Rows[0].ItemArray[12].ToString();
                ReportDocument rpt = new ReportDocument();
                rpt.Load(nf);
                rpt.SetDataSource(_dataSet1);
                crystalReportViewer1.ReportSource = rpt;
            }
        }
    }
}
