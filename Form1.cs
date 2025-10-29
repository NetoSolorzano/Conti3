using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace Conti3
{
    public partial class Form1 : Form
    {
        publicoConf conf = new publicoConf();
        public Form1()
        {
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            InitializeComponent();
            this.BackColor = Color.FromArgb(8,253,254,243);  // brillo,rojo,verde,azul rgba(253, 254, 243, 0.8)
            this.Text = "Conti3 V.3 - 2025";
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            //

        }
        public void CargaINI(Form forma)
        {
            foreach (Control oControl in forma.Controls)
            {
                if (oControl is TextBox)
                {
                    oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                    oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                }
                if (oControl is MaskedTextBox)
                {
                    oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                    oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                }
                if (oControl is Label)
                {
                    if (oControl.Name == "eti_tituloForm")
                    {
                        oControl.Font = new System.Drawing.Font(conf.nombreFont, 14);
                        oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                    }
                    else
                    {
                        oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                        oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                    }
                }
                if (oControl is CheckBox)
                {
                    oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                    oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                }
                if (oControl is RadioButton)
                {
                    oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                    oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                }
                if (oControl is ListBox)
                {
                    oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                    oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                }
                if (oControl is Panel)
                {
                    foreach (Control control in oControl.Controls)
                    {
                        if (control is TextBox)
                        {
                            control.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            control.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                        }
                        if (oControl is MaskedTextBox)
                        {
                            oControl.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            oControl.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                        }
                        if (control is Label)
                        {
                            control.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            control.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                        }
                        if (control is CheckBox)
                        {
                            control.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            control.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                        }
                        if (control is RadioButton)
                        {
                            control.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            control.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                        }
                        if (control is ListBox)
                        {
                            control.Font = new System.Drawing.Font(conf.nombreFont, conf.tamañoFont);
                            control.ForeColor = System.Drawing.Color.FromName(conf.colorFont);
                        }
                    }
                }
            }
        }       // pinta de colores al mundo!
    }
    public class generalTextBox : TextBox
    {
        publicoConf conf = new publicoConf();
        public generalTextBox()
        {
            Font = new Font(conf.nombreFont, conf.tamañoFont);
            BackColor = Color.FromName(conf.nombreFondo);
            ForeColor = Color.FromName(conf.colorFont);
            BorderStyle = BorderStyle.None;
            KeyDown += TextBox_KeyDown;
        }
        //bool allowSpace = false;
        public void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }
    public class NumericTextBox : TextBox
    {
        publicoConf conf = new publicoConf();
        
        public NumericTextBox()
        {
            Font = new Font(conf.nombreFont, conf.tamañoFont);
            BackColor = Color.FromName(conf.nombreFondo); // Color.Aqua;
            ForeColor = Color.FromName(conf.colorFont);
            BorderStyle = BorderStyle.None;
            KeyDown += TextBox_KeyDown;
        }
        bool allowSpace = false;
        // Restricts the entry of characters to digits (including hex), the negative sign,
        // the decimal point, and editing keystrokes (backspace).
        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);

            NumberFormatInfo numberFormatInfo = System.Globalization.CultureInfo.CurrentCulture.NumberFormat;
            string decimalSeparator = numberFormatInfo.NumberDecimalSeparator;
            string groupSeparator = numberFormatInfo.NumberGroupSeparator;
            string negativeSign = numberFormatInfo.NegativeSign;

            string keyInput = e.KeyChar.ToString();

            if (Char.IsDigit(e.KeyChar))
            {
                // Digits are OK
            }
            else if (keyInput.Equals(decimalSeparator) || keyInput.Equals(groupSeparator) ||
             keyInput.Equals(negativeSign))
            {
                // Decimal separator is OK
            }
            else if (e.KeyChar == '\b')
            {
                // Backspace key is OK
            }
            //    else if ((ModifierKeys & (Keys.Control | Keys.Alt)) != 0)
            //    {
            //     // Let the edit control handle control and alt key combinations
            //    }
            else if (this.allowSpace && e.KeyChar == ' ')
            {

            }
            else
            {
                // Swallow this invalid key and beep
                e.Handled = true;
                //    MessageBeep();
            }
        }
        public void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }
    public class fechaTextBox : MaskedTextBox
    {
        publicoConf conf = new publicoConf();
        public fechaTextBox()
        {
            Font = new Font(conf.nombreFont, conf.tamañoFont);
            BackColor = Color.FromName(conf.nombreFondo); // Color.Aqua;
            ForeColor = Color.FromName(conf.colorFont);
            BorderStyle = BorderStyle.None;
            KeyDown += TextBox_KeyDown;
        }
        public void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }
    public class generalEtiqueta : Label
    {
        publicoConf conf = new publicoConf();
        public generalEtiqueta()
        {
            Font = new Font(conf.nombreFont, conf.tamañoFont);
            BackColor = Color.FromArgb(conf.fondoRojoE, conf.fondoVerdeE, conf.fondoAzulE);  // FromName(conf.fondoEtiq);
            ForeColor = Color.FromName(conf.colorFont);
        }
        //bool allowSpace = false;
    }
    public class generalBoton : Button
    {
        publicoConf conf = new publicoConf();
        public generalBoton()
        {
            this.BackColor = Color.FromName(conf.colorfondoBoton);
            this.FlatStyle = FlatStyle.Popup;
            this.Text = "";
            this.Font = new Font(conf.nomFontBoton, conf.tamañoFontBoton);
        }
    }
    public class panelGeneral : Panel
    {
        publicoConf conf = new publicoConf();
        public panelGeneral()
        {
            BackColor = Color.FromName(conf.nombreFondo);
            ForeColor = Color.FromName(conf.colorFont);
            BorderStyle = BorderStyle.FixedSingle;
        }
    }
    public class radioBoton : RadioButton
    {
        publicoConf conf = new publicoConf();
        public radioBoton()
        {
            BackColor = Color.FromName(conf.nombreFondo);
            ForeColor = Color.FromName(conf.colorFont);
            Font = new Font(conf.nombreFont, conf.tamañoFont);
        }
    }
    public class selecFecha : DateTimePicker
    {
        public selecFecha()
        {
            Format = DateTimePickerFormat.Short;
            KeyDown += TextBox_KeyDown;
        }
        public void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // quitamos el sonido DING al dar enter en un textbox
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }
    public class ComboboxItem : ComboBox
    {
        public override string Text { get; set; }
        public object Value { get; set; }

        public override string ToString()
        {
            return Text;
        }
        public ComboboxItem()
        {
            DropDownStyle = ComboBoxStyle.DropDownList;
        }
    }
    public class AutoClosingMessageBox
    {
        System.Threading.Timer _timeoutTimer;
        string _caption;
        AutoClosingMessageBox(string text, string caption, int timeout)
        {
            _caption = caption;
            _timeoutTimer = new System.Threading.Timer(OnTimerElapsed,
                null, timeout, System.Threading.Timeout.Infinite);
            using (_timeoutTimer)
                MessageBox.Show(text, caption);
        }
        public static void Show(string text, string caption, int timeout)
        {
            new AutoClosingMessageBox(text, caption, timeout);
        }
        void OnTimerElapsed(object state)
        {
            IntPtr mbWnd = FindWindow("#32770", _caption); // lpClassName is #32770 for MessageBox
            if (mbWnd != IntPtr.Zero)
                SendMessage(mbWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            _timeoutTimer.Dispose();
        }
        const int WM_CLOSE = 0x0010;
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
    }
    // objetos funcionales
    public class catEgresos
    {
        public string codigo { get; set; }
        public string nombre { get; set; }  // descrizionerid
        public string largo { get; set; }   // descrizione
    }
    public class catIngresos
    {
        public string codigo { get; set; }
        public string nombre { get; set; }  // descrizionerid
        public string largo { get; set; }   // descrizione
    }
    public class monedas
    {
        public string codigo { get; set; }
        public string siglas { get; set; }
        public string nombre { get; set; }
    }
    public class cajDestino
    {
        public string codigo { get; set; }
        public string nombre { get; set; }  // descrizionerid
        public string largo { get; set; }    // descrizione
    }
    public class provees
    {
        public string codigo { get; set; }
        public string nombre { get; set; }
        public string ruc { get; set; }     // ruc del proveedor
        public string cuenta { get; set; }  // cuenta bancaria
    }
    public class montos
    {
        public decimal monOrige { get; set; }       // monto en la moneda origen
        public string codMOrige { get; set; }       // codigo de la moneda origen
        public decimal monSoles { get; set; }       // monto equivalente en soles
        public decimal tipCDol { get; set; }        // tipo de cambio dolar
        public decimal monDolar { get; set; }       // monto equivalente en dolares
        public decimal tipCOri { get; set; }        // tipo de cambio de la moneda origen
        public decimal monEuros { get; set; }       // monto equivalente en Euros
    }
    public class giroConto
    {
        public string tipodes { get; set; }         // tipo OMG o Personal       | OMG/PER
        public string ctades { get; set; }          // descrizionerid            | nombre corto cta destino
        public string codigo { get; set; }          // codigo del giroconto      | PER<id_tabla>/OMG<id_tabla>
        public string largo { get; set; }           // descrizione               | nombre largo de la cuenta
        public string idcod { get; set; }           // idcodice                  | idcodice de la cuenta
    }
    public class tipcamDia
    {
        public decimal tcD { get; set; }
        public decimal tcE { get; set; }
    }
    public class ccolores
    {
        public string Fondo_fuerte { get; set; }
        public string Fondo_suave { get; set; }
        public string Fondo_normal { get; set; }
        public string Fondo_boton_graba { get; set; }
        public string Fondo_pageFrame { get; set; }
        public string Fondo_grilla { get; set; }
        public string Grilla_fila_normal { get; set; }
        public string Grilla_fila_anulada { get; set; }
        public string Resaltado_amarillo { get; set; }
    }
}
