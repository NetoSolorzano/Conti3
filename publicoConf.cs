using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Conti3
{
    public class publicoConf
    {
        // color de fondo principal
        public int fondoPrinBrilloE = 8;  // rgba(253, 254, 243, 0.8)
        public int fondoPrinRojoE = 253;
        public int fondoPrinVerdeE = 254;
        public int fondoPriAzulE = 243;
        #region font de etiquetas y textbox
        public string nombreFont = "Verdana";
        public string nombreFondo = "White";
        public string colorFont = "Black";   // Blue
        // etiquetas          // Color.FromArgb(142,242,243,219);  // brillo,rojo,verde,azul
        public int fondoBrilloE = 142;  
        public int fondoRojoE = 242;
        public int fondoVerdeE = 243;
        public int fondoAzulE = 219;
        public int tamañoFont = 11;  // 8
        #endregion
        #region font de botones
        public string nomFontBoton = "Arial";
        public string colorFontBoton = "Black";
        public int tamañoFontBoton = 11;
        public string colorfondoBoton = "Green";
        public int fondoBrilloB = 55;
        public int fondoRojoB = 168;
        public int fondoVerdeB = 239;
        public int fondoAzulB = 131;

        #endregion
        public void sololee(Form lfrm)
        {
            foreach (Control oControls in lfrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is ComboBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is RadioButton)
                {
                    oControls.Enabled = false;
                }
                if (oControls is DateTimePicker)
                {
                    oControls.Enabled = false;
                }
                if (oControls is MaskedTextBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is GroupBox)
                {
                    oControls.Enabled = false;
                }
                if (oControls is CheckBox)
                {
                    oControls.Enabled = false;
                }
            }
        }
        public void escribe(Form efrm)
        {
            foreach (Control oControls in efrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is ComboBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is RadioButton)
                {
                    oControls.Enabled = true;
                }
                if (oControls is DateTimePicker)
                {
                    oControls.Enabled = true;
                }
                if (oControls is MaskedTextBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is GroupBox)
                {
                    oControls.Enabled = true;
                }
                if (oControls is CheckBox)
                {
                    oControls.Enabled = true;
                }
            }
        }
        public void limpiar(Form ofrm)
        {
            foreach (Control oControls in ofrm.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
            }
        }
        public void limpia_chk(Form oForm)
        {
            foreach (Control oControls in oForm.Controls)
            {
                if (oControls is CheckBox)
                {
                    CheckBox chk = oControls as CheckBox;
                    chk.Checked = false;
                }
            }
        }
        public void limpia_cmb(Form oForm)
        {
            foreach (Control oControls in oForm.Controls)
            {
                if (oControls is ComboBox)
                {
                    ComboBox cmb = oControls as ComboBox;
                    cmb.SelectedIndex = -1;
                }
            }
        }
        public void limpiapag(TabPage pag)
        {
            foreach (Control oControls in pag.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
                if (oControls is CheckBox)
                {
                    CheckBox chk = oControls as CheckBox;
                    chk.Checked = false;
                }
                if (oControls is ComboBox)
                {
                    ComboBox cmb = oControls as ComboBox;
                    cmb.SelectedIndex = -1;
                }
            }
        }
        public void limpiagbox(GroupBox gbox)
        {
            foreach (Control oControls in gbox.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
                if (oControls is CheckBox)
                {
                    CheckBox chk = oControls as CheckBox;
                    chk.Checked = false;
                }
                if (oControls is ComboBox)
                {
                    ComboBox cmb = oControls as ComboBox;
                    cmb.SelectedIndex = -1;
                }
            }
        }
        public void limpiasplit(SplitContainer split)
        {
            foreach (Control oControls in split.Panel1.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
                if (oControls is CheckBox)
                {
                    CheckBox chk = oControls as CheckBox;
                    chk.Checked = false;
                }
                if (oControls is ComboBox)
                {
                    ComboBox cmb = oControls as ComboBox;
                    cmb.SelectedIndex = -1;
                }
            }
            foreach (Control oControls in split.Panel2.Controls)
            {
                if (oControls is TextBox)
                {
                    oControls.Text = "";
                }
                if (oControls is CheckBox)
                {
                    CheckBox chk = oControls as CheckBox;
                    chk.Checked = false;
                }
                if (oControls is ComboBox)
                {
                    ComboBox cmb = oControls as ComboBox;
                    cmb.SelectedIndex = -1;
                }
            }
        }
        public string iplan()                                               // retorna la IP lan del cliente
        {
            string strHostName = System.Net.Dns.GetHostName();
            IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);
            foreach (IPAddress ipAddress in ipEntry.AddressList)
            {
                if (ipAddress.AddressFamily.ToString() == "InterNetwork")
                {
                    return ipAddress.ToString();
                }
            }
            return "--";
        }
        public string ipwan()                                               // retorna la IP wan del cliente
        {
            string externalip = "";
            try
            {
                externalip = new WebClient().DownloadString("http://icanhazip.com");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Imposible obtener dirección WAN");
            }
            return externalip;
        }

    }
}
