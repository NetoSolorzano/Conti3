using System;
using System.Text;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;
using System.Windows.Forms;
using System.Data;
using System.Globalization;
using System.Net;
using System.IO;
using System.Security.Cryptography;     
using System.Configuration;
// se agrega la siguiente referencia para enviar texto a impresora
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Conti3
{
    class libreria
    {
        string asd = Conti3.Program.vg_user;                             // usuario conectado al sistema
        //string verapp = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
        // conexion a la base de datos
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        //
        public string ult_mov(string formu, string tabla, string usuar)     // ultimo movimiento del usuario
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)   // conl.State == ConnectionState.Open
            {
                MySqlCommand ult_mov = new MySqlCommand("call actividadBd(@nfo,@tab,@usu,@uipl,@uipw)", conl);
                ult_mov.Parameters.AddWithValue("@nfo", formu);
                ult_mov.Parameters.AddWithValue("@tab", tabla);
                ult_mov.Parameters.AddWithValue("@usu", usuar);
                ult_mov.Parameters.AddWithValue("@uipl", iplan());
                ult_mov.Parameters.AddWithValue("@uipw", Conti3.Program.vg_ipwan);
                try
                {
                    ult_mov.ExecuteNonQuery();
                    retorna = "OK";
                }
                catch (MySqlException ex)
                {
                    return ex.Message;
                }
                finally { conl.Close(); }  // conl.Close()
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string md5(string password)                                  // recibe un texto password y lo encrypta con md5
        {
            System.Security.Cryptography.MD5 md5;
            md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
            Byte[] bytecoded = md5.ComputeHash(ASCIIEncoding.Default.GetBytes(password));

            return
                System.Text.RegularExpressions.Regex.Replace(BitConverter.ToString(bytecoded).ToLower(), @"-", "");
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
        public string nbname()                                              // retorna el nombre de la pc cliente
        {
            return Environment.MachineName;
        }
        public string Encrypt(string toEncrypt, bool useHashing)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            System.Configuration.AppSettingsReader settingsReader =
                                                new AppSettingsReader();
            // Get the key from config file

            //string key = (string)settingsReader.GetValue("pass",typeof(String));   // SecurityKey
            string key = "8312@Sorol";
            //System.Windows.Forms.MessageBox.Show(key);
            //If hashing use get hashcode regards to your key
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //Always release the resources and flush data
                // of the Cryptographic service provide. Best Practice

                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes.
            //We choose ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)

            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            //transform the specified region of bytes array to resultArray
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor
            tdes.Clear();
            //Return the encrypted data into unreadable string format
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
        public string Decrypt(string cipherString, bool useHashing)
        {
            byte[] keyArray;
            //get the byte code of the string

            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            System.Configuration.AppSettingsReader settingsReader =
                                                new AppSettingsReader();
            //Get your key from config file to open the lock!
            //string key = (string)settingsReader.GetValue("pass",typeof(String));   // SecurityKey
            string key = "8312@Sorol";
            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(
                                 toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
        public string mesletras(int mes)                                    // mes en letras
        {
            string mlet = "";
            switch (mes)
            {
                case 1:
                    mlet = "Enero";
                    break;
                case 2:
                    mlet = "Febrero";
                    break;
                case 3:
                    mlet = "Marzo";
                    break;
                case 4:
                    mlet = "Abril";
                    break;
                case 5:
                    mlet = "Mayo";
                    break;
                case 6:
                    mlet = "Junio";
                    break;
                case 7:
                    mlet = "Julio";
                    break;
                case 8:
                    mlet = "Agosto";
                    break;
                case 9:
                    mlet = "Setiembre";
                    break;
                case 10:
                    mlet = "Octubre";
                    break;
                case 11:
                    mlet = "Noviembre";
                    break;
                case 12:
                    mlet = "Diciembre";
                    break;
            }
            return mlet;
        }
        public string plan_excel(string formu,string campo)                 // retorna nombre plantilla excel
        {
            string valor = "";
            string consulta = "select valor from enlaces where formulario=@for and campo=@cam";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@for", formu);
                micon.Parameters.AddWithValue("@cam", campo);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            valor = dr.GetString(0).Trim();
                            dr.Close();
                        }
                    }
                    else
                    {
                        dr.Close();
                        MessageBox.Show("No se ubica la plantilla!", "Error en configuración del sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                        return null;
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en obtener constante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                    return null;
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return valor;
        }
        public string nom_prop(string np)                                   // nombre de la propiedad
        {
            string pa = "";
            string nomret = "";
            switch (np)
            {
                case "1":
                    pa = "prop01";
                    break;
                case "2":
                    pa = "prop02";
                    break;
                case "3":
                    pa = "prop03";
                    break;
                case "4":
                    pa = "prop04";
                    break;
                case "5":
                    pa = "prop05";
                    break;
                case "6":
                    pa = "prop06";
                    break;
                case "7":
                    pa = "prop07";
                    break;
                case "8":
                    pa = "prop08";
                    break;
                case "9":
                    pa = "prop09";
                    break;
                case "10":
                    pa = "prop10";
                    break;
            }
            string consulta = "select value from confmod where param=@pa";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@pa", pa);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        nomret = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en libreria: acceso a confmod");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return nomret;
        }
        public void marca_prop(bool p1,bool p2,bool p3,bool p4,bool p5,bool p6,bool p7,bool p8,bool p9,bool p10)            
        {
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State != ConnectionState.Open)
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            for(int i=1;i<11;i++)
            {
                string consulta="";
                switch(i)
                {
                    case 1:
                        if(p1==true) { consulta="update confmod set used='1' where param='prop01'"; }
                        break;
                    case 2:
                        if(p2==true) { consulta="update confmod set used='1' where param='prop02'"; }
                        break;
                    case 3:
                        if(p3==true) { consulta="update confmod set used='1' where param='prop03'"; }
                        break;
                    case 4:
                        if(p4==true) { consulta="update confmod set used='1' where param='prop04'"; }
                        break;
                    case 5:
                        if(p5==true) { consulta="update confmod set used='1' where param='prop05'"; }
                        break;
                    case 6:
                        if(p6==true) { consulta="update confmod set used='1' where param='prop06'"; }
                        break;
                    case 7:
                        if(p7==true) { consulta="update confmod set used='1' where param='prop07'"; }
                        break;
                    case 8:
                        if(p8==true) { consulta="update confmod set used='1' where param='prop08'"; }
                        break;
                    case 9:
                        if(p9==true) { consulta="update confmod set used='1' where param='prop09'"; }
                        break;
                    case 10:
                        if(p10==true) { consulta="update confmod set used='1' where param='prop10'"; }
                        break;
                }
                if (consulta != "")
                {
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    try
                    {
                        micon.ExecuteNonQuery();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error actualizando propiedad "+i.ToString().Trim(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                        return;
                    }
                }
            }
            conl.Close();
        }
        public bool valiruc(string vr,string tp)                            // validación del documento
        {   // vr = num.documento, tp=tipo documento
            int cdig = 0;
            string docu = "";
            string consulta = "select codigo,descrizionerid from desc_doc where idcodice=@tp";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@tp", tp);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    cdig = dr.GetInt16(0);
                    docu = dr.GetString(1);
                }
                else
                {
                    MessageBox.Show("Error fatal en valiruc");
                    Application.Exit();
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            bool resultado = false;
            short ultd=0;
            short [,] matris;            //dime matris(3,11)
            if(docu == "RUC")    //switch (docu)
            {
                matris = new short[3,11];
                if(vr.Length!=11)           //len(vpar1)<>11
                {
	                resultado=false;
                } 
                else
                {
                    for(short i=0;i<=10;i++)
                    {                           // for i=1 to 11
                        matris[0,i]=0;          // matris(1,i)=0
                        matris[1,i]=0;          // matris(2,i)=0
                        matris[2,i]=0;          // matris(3,i)=0
                    }                           // endfor
                    for(short i=0;i<=10;i++)      // for i=1 to 11
                    {                           // matris(1,i)=val(substr(vpar1,i,1))
                        matris[0,i]=short.Parse(vr.Substring(i,1));
                    }                           //endfor
                    matris[1,0]=5;              // matris(2,1)=5
                    matris[1,1]=4;             // matris(2,2)=4
                    matris[1,2]=3;             // matris(2,3)=3
                    matris[1,3]=2;             // matris(2,4)=2
                    matris[1,4]=7;             // matris(2,5)=7
                    matris[1,5]=6;              // matris(2,6)=6
                    matris[1,6]=5;              // matris(2,7)=5
                    matris[1,7]=4;              // matris(2,8)=4
                    matris[1,8]=3;                // matris(2,9)=3
                    matris[1,9]=2;            // matris(2,10)=2
                    matris[1,10]=0;            // matris(2,11)=0
                    for(short i=0;i<=9;i++)      // for i=1 to 10
                    {
                        matris[2,i]=(short)(matris[0,i]*matris[1,i]);    // matris(3,i)=matris(1,i)*matris(2,i)
                        matris[2,10]=(short)(matris[2,10]+matris[2,i]);    // matris(3,11)=matris(3,11)+matris(3,i)
                    }                           // endfor
                    short res1 = (short)(matris[2,10] % 11);     // res1=mod(matris(3,11),11)
                    ultd = (short)(11 - res1);            // ultd=11-res1
                    switch(ultd)
                    {
                        case 11:                        // do case
                            ultd=1;                     //case ultd=11 or ultd=1
                            break;                      // 		ultd=1
                        case 1:                         // case ultd=10 or ultd=0
                            ultd=1;                     // ultd=0
                            break;                      // endcase
                        case 10:
                            ultd=0;
                            break;
                        case 0:
                            ultd=0;
                            break;
                    }
                    if (ultd == matris[0, 10])              // if ultd=matris(1,11)
                    {                                   // else
                        resultado = true;                 // =MESSAGEBOX("RUC no Valido",0,"  Atención  ")
                    }
                    else resultado = false;             // retorno=.F.
                }
            }
            else
            {   // otros documentos sin algoritmo de validacion
                if (vr.Length == cdig)
                {
                    resultado = true;
                }
                else resultado = false;
            }
            return resultado;
        }
        public string enlaces(string formu,string campo,string param)
        {   // devuelve el valor de la tabla enlaces segun el dato buscado
            string valor = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State != ConnectionState.Open)
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            if (formu != "" && campo != "")
            {
                if(string.IsNullOrEmpty(param))
                {
                    string consulta = "select upper(valor) from enlaces where formulario=@for and campo=@cam";
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    micon.Parameters.AddWithValue("@for", formu);
                    micon.Parameters.AddWithValue("@cam", campo);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        valor = dr.GetString(0);
                    }
                    dr.Close();
                }
                else
                {
                    string consulta = "select upper(valor) from enlaces where formulario=@for and campo=@cam and param=@par";
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    micon.Parameters.AddWithValue("@for", formu);
                    micon.Parameters.AddWithValue("@cam", campo);
                    micon.Parameters.AddWithValue("@par", param);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        valor = dr.GetString(0);
                    }
                    dr.Close();
                }
            }
            conl.Close();
            return valor;
        }    // retorna valor de enlace
        public int indesc(string dato)                                      // retorna el cnt del codigo descrittive
        {
            Int16 valor = -1;
            if (dato != "")
            {
                string vista = "desc_" + dato.Substring(0, 3).ToLower();
                string consulta = "select cnt from " + vista + " where idcodice=@dat";
                MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
                conl.Open();
                if (conl.State == ConnectionState.Open)
                {
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    micon.Parameters.AddWithValue("@dat", dato);
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.Read())
                        {
                            valor = dr.GetInt16(0);
                        }
                        dr.Close();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en acceso a vistas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                    finally { conl.Close(); }
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                    Application.Exit();
                }
            }
            return valor;
        }
        public string gofirts(string tabla)                                 // devuelte el ID del primer registro de la tabla
        {
            string retorno = "";
            string consulta = "select id from " + tabla + " limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string gofirts(string tabla, string campo, string ida)       // devuelve el id del primer registro que coincide con param ID
        {   // tabla=nombre de la tabla || strinc campo=nombre campo  || ida=valor del campo
            string retorno = "";
            string consulta = "";
            consulta = "select id from " + tabla + " where " + campo + " = @ida limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@ida", ida);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string goback(string tabla,string regact)                    // devuelte el ID del registro anterior al actual
        {
            string retorno = "";
            string consulta = "";
            if (regact == "") consulta = "select id from " + tabla + " order by id desc limit 1";
            else consulta = "select id from " + tabla + " where id <= @ra order by id desc limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                if (regact != "") micon.Parameters.AddWithValue("@ra", Int32.Parse(regact) - 1);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string goback(string tabla, string regact, string campo, string ida)       // devuelte el ID del registro anterior al actual
        {
            string retorno = "";
            string consulta = "";
            if (regact != "")   //consulta = "select id from " + tabla + " order by id desc limit 1";
            {
                consulta = "select id from " + tabla + " where id <= @ra and " + campo + " = @ida order by id desc limit 1";
                MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
                conl.Open();
                if (conl.State == ConnectionState.Open)
                {
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    if (regact != "")
                    {
                        micon.Parameters.AddWithValue("@ra", Int32.Parse(regact) - 1);
                        micon.Parameters.AddWithValue("@ida", ida);
                    }
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.Read())
                        {
                            retorno = dr.GetString(0);
                        }
                        dr.Close();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message);
                        Application.Exit();
                    }
                    finally { conl.Close(); }
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                    Application.Exit();
                }
            }
            return retorno;
        }
        public string gonext(string tabla, string regact)                   // devuelte el ID del registro siguiente al actual
        {
            string retorno = "";
            string consulta = "";
            if (regact == "") consulta = "select id from " + tabla + " order by id asc limit 1";
            else consulta = "select id from " + tabla + " where id >= @ra limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                if (regact != "") micon.Parameters.AddWithValue("@ra", Int32.Parse(regact) + 1);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string gonext(string tabla, string regact, string campo, string ida)       // devuelte el ID del registro siguiente al actual
        {
            string retorno = "";
            string consulta = "";
            if (regact != "")   //consulta = "select id from " + tabla + " order by id asc limit 1";
            {
                consulta = "select id from " + tabla + " where id >= @ra and " + campo + "=@ida limit 1";
                MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
                conl.Open();
                if (conl.State == ConnectionState.Open)
                {
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    if (regact != "")
                    {
                        micon.Parameters.AddWithValue("@ra", Int32.Parse(regact) + 1);
                        micon.Parameters.AddWithValue("@ida", ida);
                    }
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.Read())
                        {
                            retorno = dr.GetString(0);
                        }
                        dr.Close();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message);
                        Application.Exit();
                    }
                    finally { conl.Close(); }
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                    Application.Exit();
                }
            }
            return retorno;
        }
        public string golast(string tabla)                                  // devuelte el ID del ultimo registro de la tabla
        {
            string retorno = "";
            string consulta = "select id from " + tabla + " order by id desc limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string golast(string tabla, string campo, string ida)        // devuelve el id del primer registro que coincide con param ID
        {   // tabla=nombre de la tabla || strinc campo=nombre campo  || ida=valor del campo
            string retorno = "";
            string consulta = "";
            consulta = "select id from " + tabla + " where " + campo + " = @ida order by id desc limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@ida", ida);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public void series(string tipo, string serie, string numero)        // actualiza correlativo de la serie
        {
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "update series set actual=@act where tipdoc=@td and serie=@se";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@act", numero);
                micon.Parameters.AddWithValue("@td", tipo);
                micon.Parameters.AddWithValue("@se", serie);
                try
                {
                    micon.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en actualización de series", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                    return;
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
        }
        public string numser(string tipo, string serie)                     // obtiene el correlativo actual
        {
            string retorno = "";
            string consulta = "select actual from series where tipdoc=@td and serie=@se";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@td", tipo);
                micon.Parameters.AddWithValue("@se", serie);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en consulta de la serie");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string corremas(string corre)                                // aumenta el correlativo en uno (1)
        {
            string retorno = "";
            if (corre == "") corre = "0";
            retorno = Right("000000" + (Int32.Parse(corre) + 1).ToString(), 7);

            return retorno;
        }
        public string serOrgDes(string tipdoc,string codanu,string codori, string coddes) // retorna serie del origen-destino
        {
            string retorno = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select serie from series where tipdoc=@doc and status<>@sta and sede=@ori and destino=@des";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@sta", codanu);
                micon.Parameters.AddWithValue("@doc", tipdoc);
                micon.Parameters.AddWithValue("@ori", codori);
                micon.Parameters.AddWithValue("@des", coddes);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorno = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en consulta de la serie");
                    Application.Exit();
                }
                finally {
                    micon.Dispose();
                    conl.Close(); 
                }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string Right(string sValue, int iMaxLength)                  // devuelve los ultimos n caractares desde la derecha
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
        }
        public string idcli(string nocl)                                    // valida cliente según el nombre dado como parametro
        {
            string retorno = "";
            string consulta = "select id from anag_cli where nombre like concat('%',@no,'%')";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@no", nocl);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (dr.Read()) retorno = dr.GetString(0);
                        dr.Close();
                    }
                    else
                    {
                        dr.Close();
                        retorno = "0";
                    }
                }
                catch(MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error!");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public Int16 matris(string codigo)                                  // devuelve el entero de la matris de usuarios
        {
            Int16 retorna = 99;
            string consulta = "select a.tipuser,a.nivel,b.cnt,c.flag1 " + 
                "from usuarios a " +
                "left join desc_tpu b on b.idcodice=a.tipuser " +
                "left join desc_nvu c on c.idcodice=a.nivel " +
                "where a.nom_user=@use";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@use", codigo);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        string conca = dr.GetString(2) + dr.GetString(3);
                        dr.Close();
                        retorna = Int16.Parse(conca);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en consulta");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string nomuser(string codigo)                                // retorna el nombre de usuario
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select nombre from usuarios where nom_user=@asd";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@asd", codigo);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string locuser(string codigo)                                // retorna el nombre del loca del usuario
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select b.descrizionerid from usuarios a " + 
                    "left join desc_loc b on b.idcodice=a.local " +
                    "where a.nom_user=@asd";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@asd", codigo);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string nomloc(string codigo)                                 // retorna nombre largo del local segun cod.usuario
        {   // parametro de entrada codigo de usuario, osea el nombre del local del usuario
            string retorna = "";
            string consulta = "select b.descrizione from usuarios a " +
                "left join desc_sds b on b.idcodice=a.local " +
                "where a.nom_user=@asd";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@asd", codigo);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string csunatloc(string codigo)                              // retorna el coódigo sunat del local
        {
            string retorna = "";
            string consulta = "select codsunat from desc_sds where idcodice=@cod";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@cod", codigo);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string nomcloc(string codigo)                                // retorna nombre largo del local segun codigo del local
        {
            string retorna = "";
            string consulta = "select a.descrizione from desc_sds a " +
                "where a.idcodice=@asd";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@asd", codigo);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string ruta(string orig, string dest)                        // retorna string ruta (origen - destino)
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select descrizione from desc_sds where idcodice=@ori";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@ori", orig);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0).Trim();
                }
                dr.Close();
                consulta = "select descrizione from desc_sds where idcodice=@des";
                micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@des", dest);
                dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = retorna + " - " + dr.GetString(0).Trim();
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string nomstat(string codest)                                // retorna el nombre del estado
        {
            string retorna = "";
            string consulta = "select descrizionerid from desc_est " +
                "where idcodice=@cod";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@cod", codest);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public decimal valtc(string fechat)                                 // retorna el tipo de cambio de la fecha en texto
        {
            decimal retorna = 0;
            string consulta = "select valor1 from cambios where fecha=@fec";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@fec", fechat);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetDecimal(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public Int16 cntdesc(string texto,string vista)                     // retorna el CNT del desc_??? segun en texto descrizionerid
        {
            Int16 retorna = -1;
            string consulta = "select cnt from desc_" + vista + " where descrizionerid like '%" + texto.Trim() + "%'";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetInt16(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public Int16 cntcodi(string codice, string vista)                   // retorna el CNT del desc_??? segun el idcodice
        {
            Int16 retorna = -1;
            string consulta = "select cnt from desc_" + vista + " where idcodice = @idc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@idc", codice);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetInt16(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string descorta(string codice, string vista)                 // retorna la descrip corta del codigo tabla descrittive
        {
            string retorna = "";
            string consulta = "select descrizionerid from desc_" + vista + " where idcodice = @idc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@idc", codice);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string serloc(string locori, string locdes, string doc)      // retorna la serie correspondiente al origen-destino y tipo de doc
        {
            string retorna = "";
            string consulta = "select serie,sede,destino from series where sede='" + locori + "' and tipdoc='" + doc + "'";   // "' and destino='" + locdes + 
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];
                    if (row[2].ToString() == "")
                    {
                        retorna = row[0].ToString();
                    }
                    else
                    {
                        if (row[2].ToString() == locdes)
                        {
                            retorna = row[0].ToString();
                        }
                    }
                }
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string serlocz(string locori, string locdes, string doc)     // retorna la serie correspondiente al origen-destino y tipo de doc y ZONA
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consta = "select valor,flag1 from desc_sds where idcodice=@ori and numero=1";
                MySqlCommand micon = new MySqlCommand(consta, conl);
                micon.Parameters.AddWithValue("@ori", locori);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    if (dr.GetInt16(0) == 1)
                    {   // una sola serie
                        retorna = dr.GetString(1);
                        dr.Close();
                    }
                    else
                    {
                        dr.Close();
                        // ,a.series
                        string query = "select a.serie from series a " +
                            "where a.sede=@ori and a.destino=@des and a.tipdoc=@tdo";
                        micon = new MySqlCommand(query, conl);
                        micon.Parameters.AddWithValue("@des", locdes);
                        micon.Parameters.AddWithValue("@ori", locori);
                        micon.Parameters.AddWithValue("@tdo", doc);
                        MySqlDataReader drq = micon.ExecuteReader();
                        if (drq.Read())
                        {
                            if (drq.GetString(0) != "")
                            {
                                retorna = drq.GetString(0);
                            }
                        }
                        drq.Close();
                        //
                        if (retorna == "")
                        {
                            string consulta = "select ifnull(a.serie,'') " +
                            "from series a " +
                            "left join desc_sds b on b.idcodice=@des " +
                            "where a.sede=@ori and a.tipdoc=@tdo and a.zona=b.codice2";
                            micon = new MySqlCommand(consulta, conl);
                            MySqlDataAdapter da = new MySqlDataAdapter(micon);
                            da.SelectCommand.Parameters.AddWithValue("@des", locdes);
                            da.SelectCommand.Parameters.AddWithValue("@ori", locori);
                            da.SelectCommand.Parameters.AddWithValue("@tdo", doc);
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                DataRow row = dt.Rows[i];
                                retorna = row[0].ToString();
                            }
                        }
                    }
                    conl.Close();
                }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string fordoc(string locori, string locdes, string doc)      // retorna el formato de impresion
        {
            string retorna = "";
            string consulta = "select format from series where sede='" + locori + "' and destino='" + locdes + "' and tipdoc='" + doc + "'";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())                      // C = chico
                {                                   // M = medio
                    retorna = dr.GetString(0);      // G = grande
                }
                else retorna = "M"; // default
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string codloc(string codigo)                                 // retorna codigo del local del usuario
        {
            string retorna = "";
            string consulta = "select a.local from usuarios a " +
                "where a.nom_user=@asd";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@asd", codigo);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) retorna = dr.GetString(0);
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string[] dirloc(string sede)                                 // retorna direccion sede desde desc_tdi
        {
            string []dires = new string[]{"","","",""};
            string consulta = "select deta1,deta2,deta3,deta4 from desc_loc where idcodice=@cod";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@cod", sede);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    dires[0] = dr.GetString(0);
                    dires[1] = dr.GetString(1);
                    dires[2] = dr.GetString(2);
                    dires[3] = dr.GetString(3);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return dires;
        }
        public string dirloca(string sede)                                  // retorna direccion desde desc_loc
        {
            string retorna = "";
            string consulta = "select concat(trim(deta1),' - ',trim(deta2),' - ',trim(deta3),' - ',trim(deta4)) " +
                "from desc_loc where idcodice=@cod";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@cod", sede);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                micon.Dispose();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string dirloca(string sede, int multi)                       // retorna la direccion de la sede en un solo string
        {
            string retorna = "";
            if (multi == 1)
            {
                string consulta = "select concat(trim(deta1),' - ',trim(deta2),' - ',trim(deta3),' - ',trim(deta4)) " +
                    "from desc_sds where idcodice=@cod";
                MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
                conl.Open();
                if (conl.State == ConnectionState.Open)
                {
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    micon.Parameters.AddWithValue("@cod", sede);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                    dr.Close();
                    conl.Close();
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                    Application.Exit();
                }
            }
            if(multi == 2)
            {
                string consulta = "select trim(deta1),trim(deta2),trim(deta3),trim(deta4) " +
                    "from desc_sds where idcodice=@cod";
                MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
                if (conl.State == ConnectionState.Open)
                {
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    micon.Parameters.AddWithValue("@cod", sede);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0) + Environment.NewLine + 
                            dr.GetString(1) + " - " + dr.GetString(2) + " - " + dr.GetString(3);
                    }
                    dr.Close();
                    conl.Close();
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                    Application.Exit();
                }
            }
            return retorna;
        }
        public string[] detserv(string orig, string dest)                   // retorna array de datos del servicio
        {
            string[] servis = new string[] { "", "", "", "", "", "", "", "" };
            string consulta = "select locori,locdes,nomserv,f1_peso,bloqueado,codigv,moneda,empaq " +
                "from maserv where concat(locori,locdes)=@rut";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@rut", orig.Trim() + dest.Trim());
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    servis[0] = dr.GetString(0);    // loc origen
                    servis[1] = dr.GetString(1);    // loc destino
                    servis[2] = dr.GetString(2);    // nombre
                    servis[3] = dr.GetString(3);    // precio kg
                    servis[4] = dr.GetString(4);    // bloqueo = "1"
                    servis[5] = dr.GetString(5);    // igv = "1"
                    servis[6] = dr.GetString(6);    // moneda
                    servis[7] = dr.GetString(7);    // empaque
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return servis;
        }
        public string glosaft(string tipdoc,string serie)                   // retorna la glosa para los docs de venta
        {
            string retorna = "";
            string consulta = "select glosaser from series where tipdoc=@tdo and serie=@ser";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@tdo", tipdoc);
                micon.Parameters.AddWithValue("@ser", serie);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string nomsn(string vista,string docu, string numero)        // retorna nombre del socio de negocio
        {
            string retorno = "";
            string consulta= "";
            if(vista=="CLI"){
                consulta = "select RazonSocial from anag_cli where IDCategoria=@vis and tipdoc=@doc and ruc=@ruc";
            }
            else consulta = "select RazonSocial from anag_for where IDCategoria=@vis and tipdoc=@doc and ruc=@ruc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@vis", vista);
                micon.Parameters.AddWithValue("@doc", docu);
                micon.Parameters.AddWithValue("@ruc", numero);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorno = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorno;
        }
        public string[] datossn(string vista, string docu, string numero)   // retorna datos del socio, (nombre,direc, etc)
        {
            string[] retorna = new string[] { "", "", "", "", "", "", "", "" };
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (vista == "CLI")
                {
                    string consulta = "select tipdoc,RUC,RazonSocial,Direcc1,Direcc2,depart,Provincia,Localidad,NumeroTel1,NumeroTel2,EMail,pais,ubigeo," +
                        "codigo,estado,idcategoria,id " +
                        "from anag_cli where tipdoc=@tdo and ruc=@ndo";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@tdo", docu);
                    micon.Parameters.AddWithValue("@ndo", numero);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorna[0] = dr.GetString("RazonSocial");
                        retorna[1] = dr.GetString("Direcc1").Trim() + " " + dr.GetString("Direcc2").Trim();
                        retorna[2] = dr.GetString("depart");
                        retorna[3] = dr.GetString("Provincia");
                        retorna[4] = dr.GetString("Localidad");
                        retorna[5] = dr.GetString("ubigeo");
                        retorna[6] = dr.GetString("NumeroTel1");
                        retorna[7] = dr.GetString("EMail");
                    }
                    dr.Dispose();
                    micon.Dispose();
                }
                if (vista == "FOR")
                {
                    // falta ...
                }
            }
            conn.Close();
            return retorna;
        }
        public string serlocs(string loca)                                  // retorna la serie del local
        {
            string retorna = "";
            string consulta = "select codigo from desc_loc where idcodice=@loc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@loc", loca);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string paisdoc(string docu)                                  // retorna el pais según el documento
        {
            string retorna = "";
            string consulta = "select deta1 from desc_doc where idcodice=@doc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@doc", docu);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string gsndoc(string docu)                                   // retorna si es nacional o extrangero el documento
        {
            string retorna = "";
            string consulta = "select codice2 from desc_doc where idcodice=@doc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@doc", docu);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string nomemple(string codi)                                 // retorna el nombre completo del empleado
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select concat(trim(nombres),' ',trim(apellidos)) from marrhhe where codemp=@cod";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@cod", codi);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string codemple(string codi)                                 // retorna codigo de empleado según su usuario
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select ifnull(codemp,' ') from marrhhe where usuario=@cod";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@cod", codi);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read())
                {
                    retorna = dr.GetString(0);
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string valientabla(string tabla, string campo, string dato)  // valida en tabla y campo, retorna 0 si no existe, 1 si
        {
            string retorna = "0";   // 0 = no existe - 1 = si existe
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                //string consulta = "select count(*) from " + tabla + " where " + campo + " = @dat";
                string consulta = "select count(*) from " + tabla + " where " + campo + " = '" + dato + "'";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                //micon.Parameters.AddWithValue("@dat", dato);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "error en buscar dato en tabla");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
            }
            return retorna;
        }
        public string valitrans(string trans,string tipo)                   // valida que la transaccion exista y este aprobada
        {
            string retorna = "0";       // 0 = no existe || >0 si existe con saldo
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);   // 
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select id,salfle,salefec,aprobado from matransac where docu=@doc";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@doc", trans);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            if (dr.GetString(3) == "")
                            {
                                retorna = "0";
                            }
                            else
                            {
                                if (tipo == "flete") retorna = dr.GetString(1);
                                if (tipo == "efect") retorna = dr.GetString(2);
                            }
                        }
                        dr.Close();
                    }
                    else
                    {
                        MessageBox.Show("No existe la transacción", "No Existe", MessageBoxButtons.OK,MessageBoxIcon.Information);
                        retorna = "0";
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "error en buscar transacción");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
            }
            return retorna;
        }
        public void creacli(string tip, string doc, string nom, string ape,
            string dir, string dis, string pro, string dep, string lgr, 
            string mon, string pai, string gru)                             // creación rapida de clientes
        {   
            string consulta = "insert into anagrafiche (" +
                "vista,docu,ruc,nombre,nombre2,pais,userc,fechc,moneda,grupo,fpago) values (" +
                "'CLI',@tip,@doc,@nom,@ape,@pai,@use,now(),@mon,@gru,@fpa)";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@tip", tip);
                micon.Parameters.AddWithValue("@doc", doc);
                micon.Parameters.AddWithValue("@nom", nom);
                micon.Parameters.AddWithValue("@ape", ape);
                micon.Parameters.AddWithValue("@pai", pai);
                micon.Parameters.AddWithValue("@use", asd);
                micon.Parameters.AddWithValue("@mon", mon);
                micon.Parameters.AddWithValue("@gru", gru);
                micon.Parameters.AddWithValue("@fpa", enlaces("frmsocneg", "cmb_fpago", "Contado"));
                try
                {
                    micon.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "NO se creo el cliente, debe hacerlo manualmente!",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    //Application.Exit();
                    return;
                }
                string lid="";
                micon = new MySqlCommand("select last_insert_id()", conl);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.Read()) lid = dr.GetString(0);
                dr.Close();
                //
                string condir = "insert into dirsocn (" +
                    "ida,direc,distr,provi,depar,locgr,userc,fechc) values (" +
                    "@ida,@dir,@dis,@pro,@dep,@loc,@use,now())";
                micon = new MySqlCommand(condir, conl);
                micon.Parameters.AddWithValue("@ida", lid);
                micon.Parameters.AddWithValue("@dir", dir);
                micon.Parameters.AddWithValue("@dis", dis);
                micon.Parameters.AddWithValue("@pro", pro);
                micon.Parameters.AddWithValue("@dep", dep);
                micon.Parameters.AddWithValue("@loc", lgr);
                micon.Parameters.AddWithValue("@use", asd);
                try
                {
                    micon.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "NO se inserto dirección del cliente");
                    //Application.Exit();
                    return;
                }
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
        }
        public string nomclie(string tip, string num)                       // retorna el nombre del cliente
        {
            string retorna = "";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    string hay = "select concat(trim(nombre),' ',trim(nombre2)) from anag_cli where ruc=@num and docu=@doc";
                    MySqlCommand micon = new MySqlCommand(hay, conn);
                    micon.Parameters.AddWithValue("@doc", tip);
                    micon.Parameters.AddWithValue("@num", num);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                    dr.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error busqueda de cliente");
                    Application.Exit();
                }
                finally { conn.Close(); }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
            }
            return retorna;
        }
        public void correo(string tipo, string numero, string mail, string socio)         // actualiza el correo del cliente/proveedor
        {
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {

                    string hay = "update anagrafiche set email=@mail where docu=@tipo and ruc=@nume and vista=@vis";
                    MySqlCommand micon = new MySqlCommand(hay, conn);
                    micon.Parameters.AddWithValue("@mail", mail);
                    micon.Parameters.AddWithValue("@tipo", tipo);
                    micon.Parameters.AddWithValue("@nume", numero);
                    if (socio == "CLI" || socio == "") micon.Parameters.AddWithValue("@vis", "CLI");
                    if (socio == "FOR") micon.Parameters.AddWithValue("@vis", "FOR");
                    micon.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error busqueda de cliente");
                    Application.Exit();
                }
                finally { conn.Close(); }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
            }
        }
        public Boolean email_bien_escrito(String email)                    // valida correo electronico, que tenga @ y . 
        {
            String expresion;
            expresion = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
            if (Regex.IsMatch(email, expresion))
            {
                if (Regex.Replace(email, expresion, String.Empty).Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        public string cero_izq(string dato, Int16 ancho)                    // retorna dato con ceros a la izq
        {
            string retorna = "";
            string yyy = retorna.PadRight(ancho, '0');
            string xxx = yyy + dato.Trim();
            retorna = xxx.Substring(xxx.Length - ancho);
            
            return retorna;
        }
        public int CentimeterToPixel(double Centimeter)                     // utilitario para calcular pixceles desde centimetros
        {
            double pixel = -1;
            //using (Graphics g = this.CreateGraphics())
            //{
            //    pixel = Centimeter * g.DpiY / 2.54d;
            //}
            pixel = (Centimeter / 2.54) * 96;
            return (int)pixel;
        }
        public string WordCut(string text, int cutOffLength, char[] separators) // retorna texto antes de un/unos separadores
        {
            cutOffLength = cutOffLength > text.Length ? text.Length : cutOffLength;
            int separatorIndex = text.Substring(0, cutOffLength).LastIndexOfAny(separators);
            if (separatorIndex > 0)
                return text.Substring(0, separatorIndex);
            return text.Substring(0, cutOffLength);
        }
        public double guiacancel(string serie, string corre)                // retorna true si las notas de la GR estan canceladas
        {
            double retorna = 0;
            string consulta = "select saldo,status from mactacte where sergr=@sgr and corgr=@cgr";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlDataAdapter da = new MySqlDataAdapter(consulta, conl);
                da.SelectCommand.Parameters.AddWithValue("@sgr", serie);
                da.SelectCommand.Parameters.AddWithValue("@cgr", corre);
                DataTable dt = new DataTable();
                da.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];
                    retorna += double.Parse(row[0].ToString());
                }
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string idcaja(string local, string fecha)                    // retorna el ID del local en la fecha
        {
            string retorna = "";
            string consulta = "select id from macajas where local=@loca and fecha=@fec";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@loca", local);
                micon.Parameters.AddWithValue("@fec", fecha.Substring(6,4)+"-"+fecha.Substring(3,2)+"-"+fecha.Substring(0,2));
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string fechCajaLoc(string loca, string estAbi)               // retorna fecha de la caja ABIERTA del local, si no esta abierta retorna en blanco
        {
            string retorna = "";    // retorna en formato ansi
            string consulta = "select DATE_FORMAT(fechope, '%Y-%m-%d') from cabccaja where loccaja=@loc and statusc=@eca";
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                    {
                        micon.Parameters.AddWithValue("@loc", loca);
                        micon.Parameters.AddWithValue("@eca", estAbi);
                        using (MySqlDataReader dr = micon.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                retorna = dr.GetString(0);
                            }
                        }
                    }
                }
            }
            return retorna;
        }
        public string idegresos(string idcaja)                              // retorna el idr de egresos segun id de caja
        {
            string retorna = "";
            string consulta = "select id from maegres where idcaja=@idc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@idc", idcaja);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string idingresos(string idcaja)                             // retorna el idr de ingresos segun el id de caja
        {
            string retorna = "";
            string consulta = "select id from maingvar where idcaja=@idc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@idc", idcaja);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string iddepositos(string idcaja)                            // retorna el idr de depositos según el id de caja
        {
            string retorna = "";
            string consulta = "select id from madepos where idcaja=@idc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@idc", idcaja);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string idtransac(string docu)                                // retorna el id de la transaccion
        {
            string retorna = "";
            string consulta = "select id from matransac where docu=@doc";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@doc", docu);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string fechaServ(string forma)                                           // retorna fecha del servidor FORMATO AAAA-MM-DD
        {
            string retorna = "";
            string consulta = "";
            if (forma == "latin") consulta = "select DATE_FORMAT(NOW(), '%d/%m/%Y') AS fecha_actual";
            if (forma == "ansi") consulta = "SELECT DATE_FORMAT(NOW(), '%Y-%m-%d') as fecha_Actual";
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                {
                    using (MySqlDataReader dr = micon.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            retorna = dr.GetString(0);
                        }
                    }
                }
            }
            return retorna;
        }
        public bool valcajero()                                             // retorna true si el usuario es cajero
        {
            bool retorna = false;
            if (matris(asd) <= 11) //  asd == "lucio"
            {
                retorna = true;
                return retorna;
            }
            else
            {
                MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
                conl.Open();
                if (conl.State == ConnectionState.Open)
                {
                    string consulta = "select cargo from marrhhe where usuario=@asd";
                    MySqlCommand micon = new MySqlCommand(consulta, conl);
                    micon.Parameters.AddWithValue("@asd", asd);
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.HasRows)
                        {
                            if (dr.Read())
                            {
                                if (dr.GetString(0) != enlaces("frmcajas", "tx_user", "Cajero"))
                                {
                                    retorna = false;
                                }
                                else retorna = true;
                            }
                            dr.Close();
                        }
                        else
                        {
                            dr.Close();
                            retorna = false;
                        }
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en consulta de cajas");
                        Application.Exit();
                    }
                    finally { conl.Close(); }
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error de conectividad");
                    Application.Exit();
                }
                return retorna;
            }
        }
        public string cargouser(string user)                                // retorna el cargo del usuario actual
        {
            string retorna = "";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                string consulta = "select cargo from marrhhe where usuario=@asd";
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@asd", user);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (dr.Read())
                        {
                            retorna = dr.GetString(0);
                        }
                        dr.Close();
                    }
                    else
                    {
                        dr.Close();
                        retorna = "";
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en consulta de cajas");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error de conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string validacaja(string loca, string estAbi)               // valida existencia de caja ABIERTA, devuelve ID de caja abierta en la fecha
        {
            string retorna = "";       // busca en la maestra de cajas buscando una que este abierta
            string consulta = "select ifnull(id,'') from cabccaja where loccaja=@loc and statusc=@sta order by id desc limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@loc", loca);
                micon.Parameters.AddWithValue("@sta", estAbi);
                try
                {
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (dr.Read()) retorna = dr.GetString(0);
                        dr.Close();
                    }
                    else
                    {
                        dr.Close();
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en busqueda de id caja");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error de conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public string lotedet(string tabla)                                 // retorna el ultimo numero de lote detraccion
        {
            string retorna = "";
            string consulta = "select lote from lotesdetra order by id desc limit 1";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                //micon.Parameters.AddWithValue("@idc", idcaja);
                MySqlDataReader dr = micon.ExecuteReader();
                if (dr.HasRows)
                {
                    if (dr.Read())
                    {
                        retorna = dr.GetString(0);
                    }
                }
                dr.Close();
                conl.Close();
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
            return retorna;
        }
        public void actmanoen(string sernot, string cornot,string sergr,string corgr,string clave)   // actualiza la GR en manoen
        {
            string consulta = "update manoen set sergre=@sgr, corgre=@cgr, keycode=@kco " +
                "where sernot=@sno and cornot=@cno";
            MySqlConnection conl = new MySqlConnection(DB_CONN_STR);
            conl.Open();
            if (conl.State == ConnectionState.Open)
            {
                MySqlCommand micon = new MySqlCommand(consulta, conl);
                micon.Parameters.AddWithValue("@sgr", sergr);
                micon.Parameters.AddWithValue("@cgr", corgr);
                micon.Parameters.AddWithValue("@sno", sernot);
                micon.Parameters.AddWithValue("@cno", cornot);
                micon.Parameters.AddWithValue("@kco", clave);
                try
                {
                    micon.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en actualizar GR en manoen");
                    Application.Exit();
                }
                finally { conl.Close(); }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor en librerías", "Error en conectividad");
                Application.Exit();
            }
        }
        public void almacen(string accion, string serie, string corre, string almac)         // actualiza almacen - descarga stock
        {   // accion = "entra" || "sale" .... ojo, todas las descargas desde notas
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (accion == "sale")
                {
                    try
                    {       // ahora mismo no hay notas en almacen, solo guias
                        string consulta = "select count(id) from almacen where grem=@scr and codloc=@loc";    //notent
                        MySqlCommand micon = new MySqlCommand(consulta, conn);
                        micon.Parameters.AddWithValue("@scr", serie + corre);
                        micon.Parameters.AddWithValue("@loc", almac);
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.HasRows)
                        {
                            if (dr.Read())
                            {
                                if (dr.GetInt16(0) > 0)
                                {
                                    dr.Close();
                                    consulta = "delete from almacen where grem=@scr and codloc=@loc";   //notent
                                    micon = new MySqlCommand(consulta, conn);
                                    micon.Parameters.AddWithValue("@scr", serie + corre);
                                    micon.Parameters.AddWithValue("@loc", almac);
                                    try
                                    {
                                        micon.ExecuteNonQuery();
                                    }
                                    catch (MySqlException ex)
                                    {
                                        MessageBox.Show(ex.Message, "Error en borrado de nota");
                                        Application.Exit();
                                        return;
                                    }
                                }
                                else dr.Close();
                            }
                            else dr.Close();
                        }
                        else
                        {
                            MessageBox.Show("No se ubica la nota en el almacén", "Error en almacén", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            dr.Close();
                            conn.Close();
                            return;
                        }
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en consulta inicial de almacen");
                        Application.Exit();
                        return;
                    }
                }
                if (accion == "entra")
                {   // reingreso por anulacion de una entrega .. o por algún caso especial
                    try
                    {
                        string sergr = "";
                        string corgr = "";
                        string canti = "";
                        string totpe = "";
                        string unida = "";
                        string descr = "";
                        string consulta = "select a.sergre,a.corgre,a.cant,a.kgtotal,b.unidad,b.descrip " +
                            "from manoen a left join detanoen b on a.sernot=b.sernot and a.cornot=b.cornot " +
                            "where a.sernot=@ser and a.cornot=@cor";
                        MySqlCommand micon = new MySqlCommand(consulta, conn);
                        micon.Parameters.AddWithValue("@ser", serie);
                        micon.Parameters.AddWithValue("@cor", corre);
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.HasRows)
                        {
                            if (dr.Read())
                            {
                                sergr = dr.GetString(0);
                                corgr = dr.GetString(1);
                                canti = dr.GetString(2);
                                totpe = dr.GetString(3);
                                unida = dr.GetString(4);
                                descr = dr.GetString(5);
                            }
                            dr.Close();
                        }
                        else
                        {
                            MessageBox.Show("No se ubica la nota", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            dr.Close();
                            conn.Close();
                            return;
                        }
                        //
                        consulta = "insert into almacen (codloc,feching,corring,notent,grem,cant,unidad,descr,tiping,userc,fechc) " +
                            "values (@loc,@fec,@cor,@not,@gre,@can,@uni,@des,@tip,@asd,now())";
                        micon = new MySqlCommand(consulta, conn);
                        micon.Parameters.AddWithValue("@loc", almac);
                        micon.Parameters.AddWithValue("@fec", DateTime.Now.ToString("yyyy-MM-dd"));
                        micon.Parameters.AddWithValue("@cor", "R");
                        micon.Parameters.AddWithValue("@not", serie + corre);
                        micon.Parameters.AddWithValue("@gre", sergr + corgr);
                        micon.Parameters.AddWithValue("@can", canti);
                        micon.Parameters.AddWithValue("@uni", unida);
                        micon.Parameters.AddWithValue("@des", descr);
                        micon.Parameters.AddWithValue("@tip", "R");
                        micon.Parameters.AddWithValue("@asd", asd);
                        micon.ExecuteNonQuery();
                        dr.Close();
                        conn.Close();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en obtener/reingresar la nota");
                        Application.Exit();
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Se perdió conexión con el servidor","Error de conexión");
                Application.Exit();
            }
        }
        public bool validac(string tabla, string campo, string valor)       // retorna true si el valor del campo existe en la tabla
        {
            bool retorna=false;
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string consulta = "select count(*) from " + tabla + " where " + campo + "='" + valor + "'";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                MySqlDataReader dr =  micon.ExecuteReader();
                if (dr.Read())
                {
                    if (dr.GetInt16(0) == 1) retorna = true;
                    else retorna = false;
                }
                dr.Close();
                conn.Close();
            }
            else
            {
                MessageBox.Show("No se pudo conectar al servidor", "Error de red");
                Application.Exit();
                retorna = false;
            }
            return retorna;
        }
        public bool actuac(string tabla, string campo, string valor, string id)       // retorna true si actualizó
        {
            bool retorna = false;
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    string consulta = "update " + tabla + " set " + campo + "='" + valor + "' where id = '" + id + "'";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.ExecuteNonQuery();
                    retorna = true;
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "No se actualizó el campo");
                    retorna = false;
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("No se pudo conectar al servidor", "Error de red");
                Application.Exit();
                retorna = false;
            }
            return retorna;
        }
        public bool validaub(string tabla, string campo, string valor, string codde, string codpr, string coddi)
        {
            bool retorna = false;
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    string consulta = "";
                    if (codde != "00" && codpr == "00" && coddi == "00")    // valida departamento
                    {
                        consulta = "select count(*) from " + tabla + " where " + campo + "='" + valor + "' and provin='00' and distri='00'";
                    }
                    if(codde != "00" && codpr != "00" && coddi == "00")
                    {
                        consulta = "select count(*) from " + tabla + " where " + campo + "='" + valor + "' and depart='" + codde + "' and distri='00'";
                    }
                    if(codde != "00" && codpr != "00" && coddi != "00")
                    {
                        consulta = "select count(*) from " + tabla + " where " + campo + "='" + valor + "' and depart='" + codde + "' and provin='" + codpr + "'";
                    }
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        if (dr.GetInt16(0) == 1) retorna = true;
                        else retorna = false;
                    }
                    dr.Close();
                    conn.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "No se genero consulta");
                    retorna = false;
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("No se pudo conectar al servidor", "Error de red");
                Application.Exit();
                retorna = false;
            }
            return retorna;
        }
        public string estcont(string cont)                                  // recalcula, actualiza y retorna el estado del contrato
        {
            string retorna = "";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string asalc = "sal_cont";
                MySqlCommand misp = new MySqlCommand(asalc,conn);
                misp.CommandType = CommandType.StoredProcedure;
                misp.CommandTimeout = 100;
                misp.Parameters.AddWithValue("@cont", cont);
                //MySqlParameter reto = misp.Parameters.Add("@estad", MySqlDbType.VarChar);
                //reto.Direction = ParameterDirection.Output;
                misp.ExecuteNonQuery();
                //
                string compa = "act_cont";
                misp = new MySqlCommand(compa, conn);
                misp.CommandType = CommandType.StoredProcedure;
                misp.CommandTimeout = 300;
                misp.Parameters.AddWithValue("@cont", cont);
                MySqlParameter reto = misp.Parameters.Add("@estad", MySqlDbType.VarChar);
                reto.Direction = ParameterDirection.Output;
                misp.ExecuteNonQuery();
                retorna = reto.Value.ToString();
            }
            else
            {
                MessageBox.Show("No se pudo conectar al servidor", "Error de red");
                Application.Exit();
            }
            return retorna;
        }
        public string funlet(int num)                                       // devuelve la letra del alfabeto acorde al numero
        {
            string retorna = "";
            string[] tabla = new string[] { "_", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "W", "X", "Y", "Z" };
            retorna = tabla[num];
            return retorna;
        }
        public string[] retDPDubigeo(string ubigeo)                         // retorna [] dep, pro y distrito del cod.ugigeo
        {
            string[] retorna = { "","",""};
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string aa = "SELECT nombre FROM ubigeos WHERE concat(depart, provin, distri)=concat(LEFT(@ubg, 2),'0000') " +
                    "union all " +
                    "SELECT nombre FROM ubigeos WHERE concat(depart, provin, distri)=concat(LEFT(@ubg, 4),'00') " +
                    "union all " +
                    "SELECT nombre FROM ubigeos WHERE concat(depart, provin, distri)=@ubg";
                MySqlCommand micon = new MySqlCommand(aa, conn);
                micon.Parameters.AddWithValue("@ubg", ubigeo);
                MySqlDataReader dr = micon.ExecuteReader();
                int i = 0;
                while (dr.Read())
                {
                    retorna[i] = dr.GetString(0);
                    i++;
                }
                dr.Dispose();
                micon.Dispose();
                conn.Close();
            }
            return retorna;
        }
        public string retCodubigeo(string nomD, string nomP, string codUb)  // retorna codigo ubigeo del nombre
        {
            string retorna="";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                if (nomD != "" && nomP == "" && codUb == "")
                {
                    string consulta = "select depart from ubigeos where trim(nombre)=@dep and depart<>'00' and provin='00' and distri='00'";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@dep", nomD.Trim());
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.HasRows == true)
                        {
                            while (dr.Read())
                            {
                                retorna = dr.GetString(0).Trim();
                            }
                        }
                        dr.Close();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en obtener codigo de departamento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }
                if (nomD == "" && nomP != "" && codUb != "")
                {
                    string consulta = "select provin from ubigeos where trim(nombre)=@prov and depart=@dep and provin<>'00' and distri='00'";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@dep", codUb.Substring(0, 2));
                    micon.Parameters.AddWithValue("@prov", nomP.Trim());
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.HasRows == true)
                        {
                            while (dr.Read())
                            {
                                if (codUb.Trim().Length == 6) retorna = codUb.Substring(0, 2) + dr.GetString(0).Trim() + codUb.Substring(4, 2);
                                if (codUb.Trim().Length < 6) retorna = codUb.Substring(0, 2) + dr.GetString(0).Trim();
                            }
                        }
                        dr.Close();
                        micon.Dispose();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en obtener codigo de provincia", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }
                if (nomD != "" && nomP == "" && codUb != "")
                {
                    string consulta = "select distri from ubigeos where trim(nombre)=@dist and depart=@dep and provin=@prov and distri<>'00'";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@dep", codUb.Substring(0, 2));
                    micon.Parameters.AddWithValue("@prov", (codUb.Length > 2) ? codUb.Substring(2, 2) : "  ");
                    micon.Parameters.AddWithValue("@dist", nomD.Trim());
                    try
                    {
                        MySqlDataReader dr = micon.ExecuteReader();
                        if (dr.HasRows == true)
                        {
                            while (dr.Read())
                            {
                                if (codUb.Trim().Length >= 4) retorna = codUb.Trim().Substring(0, 4) + dr.GetString(0).Trim();
                            }
                        }
                        dr.Close();
                        micon.Dispose();
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show(ex.Message, "Error en obtener codigo de distrito", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("No se puede conectar al servidor!", "Error de conectividad", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            return retorna;
        }
        public string[] conectorSolorsoft(string cual, string numDoc)
        {
            string[] retorna = new string[8];
            retorna[0]="";retorna[1] = ""; retorna[2] = ""; retorna[3] = "";
            retorna[4] = ""; retorna[5] = ""; retorna[6] = ""; retorna[7] = "";
            string dirLoc = Directory.GetCurrentDirectory() + @"\conectores\";
            if (cual == "RUC")
            {
                ProcessStartInfo start = new ProcessStartInfo();
                //start.FileName = dirLoc + "ConectorSolorsoft.exe";
                start.FileName = dirLoc + "conectorJson.exe";
                start.UseShellExecute = false;
                start.RedirectStandardOutput = true;
                start.Arguments = numDoc + " " + "0";  // 2 parametros, numero a buscar, tipo de doc (0=ruc|1=dni)
                using (Process proceso = Process.Start(start))
                {

                    using (StreamReader reader = proceso.StandardOutput)
                    {
                        string result = reader.ReadToEnd();
                        string[] datos = result.Split('|');

                        retorna[0] = datos[1].Replace("?","Ñ");         // razon social
                        retorna[1] = datos[4];                          // ubigeo
                        retorna[2] = datos[5];                          // direccion
                        retorna[3] = datos[8];                          // departamento
                        retorna[4] = datos[7];                          // provincia
                        retorna[5] = datos[6];                          // distrito
                        retorna[6] = datos[2];                          // estado del contrib.
                        retorna[7] = datos[3];                          // condicion domicilio
                    }
                }
            }
            if (cual == "DNI")
            {
                string nomCon = "";
                using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                {
                    conn.Open();
                    string abre = "select valor from enlaces where formulario='main' and campo='conector'";
                    using (MySqlCommand micon = new MySqlCommand(abre, conn))
                    {
                        using (MySqlDataReader dr = micon.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                nomCon = dr.GetString(0);
                            }
                        }
                    }
                }
                if (nomCon == "conectorc.exe")
                {
                    int limite = 30000;
                    ProcessStartInfo start = new ProcessStartInfo();
                    start.UseShellExecute = false;
                    start.RedirectStandardOutput = true;
                    start.FileName = dirLoc + "conectorc.exe";
                    start.Arguments = "D " + numDoc;  // 2 parametros, "D" y dni a buscar
                    Process p = Process.Start(start);
                    p.WaitForInputIdle();
                    p.WaitForExit(limite);
                    if (p.HasExited == false)
                    {
                        if (p.Responding)
                        {
                            p.CloseMainWindow();
                        }
                        else
                        {
                            p.Kill();
                        }
                    }
                    else
                    {
                        // leemos el .config del conectorc
                        ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
                        configMap.ExeConfigFilename = dirLoc + "conectorc.exe.config";
                        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
                        retorna[0] = config.AppSettings.Settings["rnombre"].Value + " " +
                            config.AppSettings.Settings["rapater"].Value + " " +
                            config.AppSettings.Settings["ramater"].Value;   // textBox9.Text NOMBRE
                        retorna[1] = config.AppSettings.Settings["rnumero"].Value;  // textBox3.Text NUMERO DNI
                        //
                        config.AppSettings.Settings["rnombre"].Value = "";
                        config.AppSettings.Settings["rapater"].Value = "";
                        config.AppSettings.Settings["ramater"].Value = "";
                        config.Save();
                        p.Close();
                    }
                }
                if (nomCon == "conectorJson.exe")
                {
                    ProcessStartInfo start = new ProcessStartInfo();
                    start.FileName = dirLoc + "conectorJson.exe";
                    start.UseShellExecute = false;
                    start.RedirectStandardOutput = true;
                    start.Arguments = numDoc + " " + "1";  // 2 parametros, numero a buscar, tipo de doc (0=ruc|1=dni)
                    using (Process proceso = Process.Start(start))
                    {
                        using (StreamReader reader = proceso.StandardOutput)
                        {
                            string result = reader.ReadToEnd();
                            if (string.IsNullOrEmpty(result))
                            {
                                retorna[1] = "";
                                retorna[0] = "";
                            }
                            else
                            {
                                string[] datos = result.Split('|');
                                retorna[1] = datos[0];
                                retorna[0] = datos[1].Replace("?", "Ñ");
                            }
                        }
                    }
                }
            }
            return retorna;
        }
        public void messagebox(string texto)                                // messagebox de alerta sobre variables
        {
            MessageBox.Show(texto, "FALTA CONFIGURAR", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        public bool procConn(MySqlConnection con)                           // conexión a la base de datos
        {
            bool retorna = false;
            while (con.State != ConnectionState.Open)
            {
                try
                {
                    con.Open();
                    if (con.State == ConnectionState.Open)
                    {
                        retorna = true;
                        break;
                    }
                }
                catch (MySqlException ex)
                {   // con.ConnectionString + Environment.NewLine +
                    var xx = MessageBox.Show(ex.Message + Environment.NewLine + 
                        "Tenemos un problema con la conexión al servidor" + Environment.NewLine +
                        "No se puede conectar o esta tardando demasiado" + Environment.NewLine +
                        "DESEA SEGUIR ESPERANDO ....?","Error de conectividad",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                    if (xx == DialogResult.No)
                    {
                        retorna = false;
                        break;
                    }
                }
            }
            return retorna;
        }
        public bool valcaja(string codcaj)              // pasa naaaa
        {
            bool retorna = false;
            // mejor no ... en cada form donde tenga que ver la caja debe jalarse la caja actual y su estado, todo esta
            // lógica debe hacerce en el jalainfo .... 03/12/2020
            return retorna;
        }
        public bool DirectoryVisible(string path)                           // valida existencia y acceso al directorio
        {
            try
            {
                Directory.GetAccessControl(path);
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public class ComboItem
    {
        public string Text { get; set; }
        public object Value { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }

    #region Clase para enviar a imprsora texto plano
    public class RawPrinterHelper
    {
        // Structure and API declarions:
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }
        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            Int32 dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.

            di.pDocName = "Conti3 C#.NET Documento RAW";
            di.pDataType = "RAW";

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        public static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            Int32 dwCount;
            // How many characters are in the string?
            dwCount = szString.Length;
            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }
    }
    #endregion
}
