using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Interfaz_MADEGOM.Functions
{
    class TFunctions
    {
        //Funcion registra log
        public void AddLog(String Mensaje)
        {
            StreamWriter Arch;
            //Exe: String := 
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String NomArch;
            String NomArchB;
            NomArch = "\\VDLog_" + String.Format("{0:yyyy-MM-dd}", DateTime.Now) + ".log";
            Arch = new StreamWriter(sPath + NomArch, true);
            NomArchB = sPath + "\\VDLog_" + String.Format("{0:yyyy-MM-dd}", DateTime.Now.AddDays(-1)) + ".log";
            //Elimina archivo del dia anterior
            //if (System.IO.File.Exists(NomArchB))
            //    System.IO.File.Delete(NomArchB);

            try
            {
                Arch.WriteLine(String.Format("{0:dd-MM-yyyy HH:mm:ss}", DateTime.Now) + " " + Mensaje);
            }
            finally
            {
                Arch.Flush();
                Arch.Close();
            }
        }

        public String[] IniciaArchivos()
        {
            XmlDocument xDoc;
            XmlNodeList Configuracion;
            XmlNodeList lista;
            TFunctions Func;
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String[] arr;
            Int32 i;

            try 
            {
                xDoc = new XmlDocument();

                xDoc.Load(sPath + "\\Config.xml");
                Configuracion = xDoc.GetElementsByTagName("Configuracion");
                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("Archivo");
                arr = new System.String[lista.Count];
                i = 0;
        
                foreach (XmlElement nodo in lista)
                {
                    //var nArchivos := nodo.GetElementsByTagName('Archivo');
                    //arr[i] := system.String(nArchivos[i].InnerText);
                    arr[i] = (System.String)(nodo.InnerText);
                    i = i + 1;
                }
        
                return arr;
            }
            catch (Exception w)
            {
                Func = new TFunctions();
                Func.AddLog("IniciaArchivos: " + w.Message + " ** Trace: " + w.StackTrace);
                return null;
            }
        }

        public String PathArchivos()
        {
            XmlDocument xDoc;
            XmlNodeList Configuracion;
            XmlNodeList lista;
            TFunctions Func;
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String _result = "";

            try 
            {
                xDoc = new XmlDocument();
                xDoc.Load(sPath + "\\Config.xml");
                Configuracion = xDoc.GetElementsByTagName("Configuracion");
                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("PathArchivos");
        
                foreach (XmlElement nodo in lista)
                {
                    var nArchivos = nodo.GetElementsByTagName("PathArchivo");
                    _result = (String)(nArchivos[0].InnerText);
                }
        
                return _result;
            }
            catch (Exception w)
            {
                Func = new TFunctions();
                Func.AddLog("PathArchivos: " + w.Message + " ** Trace: " + w.StackTrace);
                return "";
            }
        }

        public String PathArchivosResp()
        {
            XmlDocument xDoc;
            XmlNodeList Configuracion;
            XmlNodeList lista;
            TFunctions Func;
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String _result = "";

            try
            {
                xDoc = new XmlDocument();
                xDoc.Load(sPath + "\\Config.xml");
                Configuracion = xDoc.GetElementsByTagName("Configuracion");
                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("PathArchivos");

                foreach (XmlElement nodo in lista)
                {
                    var nArchivos = nodo.GetElementsByTagName("PathArchivoResp");
                    _result = (String)(nArchivos[0].InnerText);
                }

                return _result;
            }
            catch (Exception w)
            {
                Func = new TFunctions();
                Func.AddLog("PathArchivosResp: " + w.Message + " ** Trace: " + w.StackTrace);
                return "";
            }
        }


        public String DatosSFTP(String Valor)
        {
            XmlDocument xDoc;
            XmlNodeList Configuracion;
            XmlNodeList lista;
            TFunctions Func;
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String _result = "";

            try
            {
                xDoc = new XmlDocument();
                xDoc.Load(sPath + "\\Config.xml");
                Configuracion = xDoc.GetElementsByTagName("Configuracion");
                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("DatosSFTP");

                foreach (XmlElement nodo in lista)
                {
                    var nArchivos = nodo.GetElementsByTagName(Valor);
                    _result = (String)(nArchivos[0].InnerText);
                }

                return _result;
            }
            catch (Exception w)
            {
                Func = new TFunctions();
                Func.AddLog("DatosSFTP: " + w.Message + " ** Trace: " + w.StackTrace);
                return "";
            }
        }

        public void AddLogSap(SAPbobsCOM.Company oCompany, String Comentario, String Archivo, double Cantidad, String Code, String Razon, String Ref1,
                               String Ref2, String Ref4, String Orden, String ItemCode, String ASN, String LPN, String Carton )
        {

            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            try
            {
                oCompanyService = oCompany.GetCompanyService();
                // Get GeneralService (oCmpSrv is the CompanyService)
                oGeneralService = oCompanyService.GetGeneralService("LOGSAP");
                // Create data for new row in main UDO
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("U_Archivo", Archivo);
                oGeneralData.SetProperty("U_Cantidad", Cantidad);
                oGeneralData.SetProperty("U_Code", Code);
                oGeneralData.SetProperty("U_Razon", Razon);
                oGeneralData.SetProperty("U_Ref1", Ref1);
                oGeneralData.SetProperty("U_Ref2", Ref2);
                oGeneralData.SetProperty("U_Ref4", Ref4);
                oGeneralData.SetProperty("U_Orden", Orden);
                oGeneralData.SetProperty("U_ItemCode", ItemCode);
                oGeneralData.SetProperty("U_ASN", ASN);
                oGeneralData.SetProperty("U_LPN", LPN);
                oGeneralData.SetProperty("U_Carton", Carton);
                oGeneralData.SetProperty("U_Comentario", Comentario);
                oGeneralParams = oGeneralService.Add(oGeneralData); 
            }
            catch (Exception ex)
            {
                AddLog(ex.Message + "Error Exception");   
            } 
        }

        public void insertarLotesBDAux(string codigo, int cantidad, string lote, DateTime timeExp, string tipoDoc, int fkDocumentSAP, string carton)
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            string sql = null;
            SqlDataReader dataReader;

            string servidor = obtenerDatosXMLServidorSAP("Servidor");
            string usuarioSQL = obtenerDatosXMLServidorSAP("UsuarioSQL");
            string passwordSQL = obtenerDatosXMLServidorSAP("PasswordSQL");

            connetionString = "Data Source={0};Initial Catalog=HEVEA_LOGISTICA;User ID={1};Password={2}";
            connetionString = String.Format(connetionString, servidor, usuarioSQL, passwordSQL);
            
            //connetionString = "Data Source=SQL-VD;Initial Catalog=HEVEA_LOGISTICA;User ID=sa;Password=SAPB1Admin";
            sql = "INSERT INTO [dbo].[LOTESWMS]([Codigo],[Cantidad],[Lote],[Vencimiento],[TipoDoc],[FkDocumenSap],[Carton]) VALUES ('{0}',{1},'{2}','{3}','{4}',{5},'{6}')";
            sql = String.Format(sql, codigo, cantidad, lote, timeExp, tipoDoc, fkDocumentSAP,carton);
            connection = new SqlConnection(connetionString);
            try
            {
                connection.Open();
                command = new SqlCommand(sql, connection);
                dataReader = command.ExecuteReader();
                dataReader.Close();
                command.Dispose();
                connection.Close();
            }
            catch (Exception ex)
            {

                AddLog("Error con conexion BD HEVEA_LOGISTICA " + ex.Message);
            }

        }


        public string obtenerDatosXMLServidorSAP(string campo)
        {
            XmlDocument xDoc;
            XmlNodeList Configuracion;
            XmlNodeList lista;
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            string result="";

            try
            {
                xDoc = new XmlDocument();
                xDoc.Load(sPath + "\\Config.xml");

                Configuracion = xDoc.GetElementsByTagName("Configuracion");
                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("ServidorSAP");

                foreach (XmlElement nodo in lista)
                {
                    var i = 0;
                    var nServidor = nodo.GetElementsByTagName(campo);
                    result = (System.String)(nServidor[i].InnerText);
                   
                }
                return result;
            }
            catch (Exception w)
            {
                AddLog("ConectarBase: " + w.Message + " ** Trace: " + w.StackTrace);
                return "";
            }

        }





    }
}
