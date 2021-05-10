using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Net;
using System.IO;
using System.Globalization;
using VisualD.Core;
using VisualD.GlobalVid;
using Newtonsoft.Json.Serialization;
using System.Reflection;
using Newtonsoft.Json;
using SAPbobsCOM;
using System.Data.SqlClient;
using System.Data;

namespace SBO_VID_Currency
{
    public class TasaCambioSBO
    {
        public DateTime Fecha;
        public Decimal UF;
        public Decimal Dolar;
        public Decimal Euro;
    }

    public class CurrencyRates
    {
        public string Moneda { get; set; }
        public Dictionary<Decimal, DateTime> Rates { get; set; }
    }
    
    sealed class SBOControl
    {
        // Log Property
        private Logs.Logger FLog;
        public Logs.Logger oLog
        {
            get { return this.FLog; }
            set { this.FLog = value; }
        }

        //Methods
        private void GetParameters()
        {
            globals.CompanyList.Clear();
            string[] separador = new string[] { ";" };
            string[] lista = SBO_VID_Currency.Properties.Settings.Default.Companies.Split(separador, StringSplitOptions.RemoveEmptyEntries);
            foreach (String s in lista)
                globals.CompanyList.Add(s);

            globals.SBOUser = SBO_VID_Currency.Properties.Settings.Default.SBOUserName;
            globals.SBOPass = Crypt.Decrypt(SBO_VID_Currency.Properties.Settings.Default.SBOPassword, "VIDCurrency0-09");

            globals.TipoCambioDirecto = true;
            if (SBO_VID_Currency.Properties.Settings.Default.TipoCambio == 1)
                globals.TipoCambioDirecto = false;

            globals.DolarObs = SBO_VID_Currency.Properties.Settings.Default.Dolar;
            globals.Euro = SBO_VID_Currency.Properties.Settings.Default.Euro;
            globals.UF = SBO_VID_Currency.Properties.Settings.Default.UF;
            globals.conexionSAP =   bool.Parse(SBO_VID_Currency.Properties.Settings.Default.conexionSAP);
            globals.SQLType = SBO_VID_Currency.Properties.Settings.Default.SQLType;
            globals.server = SBO_VID_Currency.Properties.Settings.Default.SBOServer;
            globals.DBUser = SBO_VID_Currency.Properties.Settings.Default.DBUserName;
            globals.DBPass = Crypt.Decrypt(SBO_VID_Currency.Properties.Settings.Default.DBPassword, "VIDCurrency0-09");
                

            if (!globals.TipoCambioDirecto)
            {
                globals.DolarObsI = SBO_VID_Currency.Properties.Settings.Default.DolarI;
                globals.EuroI = SBO_VID_Currency.Properties.Settings.Default.EuroI;
                globals.UFI = SBO_VID_Currency.Properties.Settings.Default.UFI;
            }
        }

        private static T _download_serialized_json_data<T>(string url) where T : new()
        {
            using (var w = new WebClient())
            {
                var json_data = string.Empty;
                // attempt to download JSON data as a string
                try
                {
                    json_data = w.DownloadString(url);
                }
                catch (Exception) { }
                // if string with JSON data is not empty, deserialize it to class and return its instance 
                return !string.IsNullOrEmpty(json_data) ?  Newtonsoft.Json.JsonConvert.DeserializeObject<T>(json_data) : new T();
            }
        }


        public static double restGET(string URL, string Data)
        {
            try
            {
                var request = WebRequest.Create(URL);

                request.ContentType = "application/json";
                request.Method = "GET";
                request.Timeout = 45000;
                var type = request.GetType();
                var currentMethod = type.GetProperty("CurrentMethod", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(request);

                var methodType = currentMethod.GetType();
                methodType.GetField("ContentBodyNotAllowed", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(currentMethod, false);

                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    streamWriter.Write(Data);
                }

                var response = (HttpWebResponse)request.GetResponse();

                if (((HttpWebResponse)(response)).StatusCode != HttpStatusCode.OK)
                {
                    Console.WriteLine("error");
                }

                using (Stream strReader = response.GetResponseStream())
                {
                    if (strReader == null) 
                        return -1;
                    using (StreamReader objReader = new StreamReader(strReader))
                    {
                        string responseBody = objReader.ReadToEnd();

                        Result result;

                        result = JsonConvert.DeserializeObject<Result>(responseBody);
                        Console.WriteLine(responseBody);
                        Console.WriteLine(result.valor);
                        return result.valor;
                    }
                }
            }
            catch (WebException ex)
            {
                throw new Exception("Error sendRequest " + ex.Message);
            }
        }

        public static string restGETWhitParameter(string URL, string parameter)
        {
            try
            {
                var request = WebRequest.Create(URL + parameter);
                request.Method = "GET";
                request.Timeout = 15000;

                var response = (HttpWebResponse)request.GetResponse();

                if (((HttpWebResponse)(response)).StatusCode != HttpStatusCode.OK)
                {
                    Console.WriteLine("error");
                }

                using (Stream strReader = response.GetResponseStream())
                {
                    if (strReader == null)
                        return"";
                    using (StreamReader objReader = new StreamReader(strReader))
                    {
                        string responseString  = objReader.ReadToEnd();
                        return responseString;
                    }
                }
            }
            catch (WebException ex)
            {
                throw new Exception("Error sendRequest " + ex.Message);
            }
        }


        
        public void Doit(ref int nErr, ref string sErr)
        {
            SAPbobsCOM.Company oCompany = null; // The company object
            SqlConnection cnn = null; 
            int i;

            try
            {
                GetParameters();

                if (globals.conexionSAP) //trabajando con DIAPI
                {
                    if (oCompany == null)
                        oCompany = new SAPbobsCOM.Company();

                    for (i = 0; i < globals.CompanyList.Count; i++)
                    {
                        try 
                        {
                            if (!oCompany.Connected)
                            {
                                ConnectSBO(ref oCompany, ref nErr, ref sErr, (String)globals.CompanyList[i]);
                                if (nErr != 0)
                                {
                                    oLog.LogMsg(sErr, "F", "E");
                                    oCompany.Disconnect();
                                    //return;
                                }
                            }
                            if (oCompany.Connected)
                            {
                                processCurrency(oCompany);
                                oCompany.Disconnect();
                            }
                        }
                        catch (Exception e)
                        {
                            oLog.LogMsg("Error en ejecución de servicio  en empresa " +  (String)globals.CompanyList[i] + " Excepcion: "+ e.Message, "F", "E");
                            if (oCompany.Connected)
                                oCompany.Disconnect();
                        }
                    }
                }
                else // conexion DB
                {
                    if (cnn == null)
                        cnn = new SqlConnection();

                    for (i = 0; i < globals.CompanyList.Count; i++)
                    {
                        try
                        {
                            conectarDBSQL((String)globals.CompanyList[i], ref sErr, ref nErr, ref cnn);
                            if (nErr != 0)
                            {
                                oLog.LogMsg(sErr, "F", "E");
                                cnn.Close();
                                //return;
                            }
                            if (cnn.State == System.Data.ConnectionState.Open)
                            {
                                processCurrencyDB(cnn);
                                cnn.Close();
                            }
                        }
                        catch (Exception e)
                        {
                            oLog.LogMsg("Error en ejecución de servicio  en empresa " + (String)globals.CompanyList[i] + " Excepcion: " + e.Message, "F", "E");
                            if (cnn.State == System.Data.ConnectionState.Open)
                                cnn.Close();
                        }
          
                    }
                }
                oLog.LogMsg("Proceso finalizado " , "F", "I");
                System.Threading.Thread.Sleep(8000);
                System.Environment.Exit(0);
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error en ejecución de App " + e.Message, "F", "E");
                if (oCompany.Connected)
                    oCompany.Disconnect();
            }
        }

        private void conectarDBSQL(string BD, ref string sErr, ref int nErr, ref SqlConnection cnn)
        {
            //SqlConnection cnn = null;
            string s = "Password={0};Persist Security Info=True;User ID={1};Initial Catalog={2};Data Source={3}";
            VisualD.Core.TVisualDCore Core;
            sErr = "";
            nErr = 0;

            try 
            {
                s = String.Format(s, globals.DBPass, globals.DBUser, BD, globals.server);
                oLog.LogMsg("try connect", "F", "D");
                cnn = new SqlConnection(s);
                cnn.Open();

                s = "select InstallNo, GetDate() Fec from cinf";
                var command = new SqlCommand(s, cnn);
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string InstallId = ((System.String)reader.GetValue(0)).Trim();
                        DateTime Fec = ((System.DateTime)reader.GetValue(1));
                        TGlobalAddOnOptions FOpcionAddOn = new TGlobalAddOnOptions();
                        List<String> FAddonsList = new List<String>();
                        string oK = "C9A80FB4-D8C3-11E0-AEAD-7E944824019B";
                        oLog.LogMsg("Before accept", "F", "D");
                        Core = new VisualD.Core.TVisualDCore(InstallId, Fec, ref FOpcionAddOn, ref oK, FAddonsList);
                        if (Core.ExitApp)
                        {
                            sErr = "Licencia " + Core.Error;
                            nErr = -2;
                            cnn.Close();
                        }
                        else if (oK != "E6DCA176-D8C3-11E0-93B0-86944824019B")
                        {
                            sErr = "VisualDCore enlace 2 inválido.";
                            nErr = -3;
                            cnn.Close();
                        }
                        else
                            sErr = "OK";

                        if (Core.EsDEmo)
                            oLog.LogMsg("Licencia de demostración " + Core.AddOnName, "F", "E");

                        oLog.LogMsg("after accept", "F", "D");

                        oLog.LogMsg("Conexión a SBO: " + BD + " - " + sErr, "F", "E");

                    }
                    else
                    {
                        sErr = "Installation Number no encontrado";
                        nErr = -1;
                        return;
                    }
                    
                }
            }
            catch (Exception e)
            {
                oLog.LogMsg("Conexión BD: " + BD + " " + sErr + "Exeption: "+ e.Message , "F", "E");
            }
        }

        private void ConnectSBO(ref SAPbobsCOM.Company oCompany, ref int nErr, ref string sErr, string CompanyName)
        {
            VisualD.Core.TVisualDCore Core;

            sErr = "";
            nErr = 0;

            oLog.LogMsg("ConnectSBO start", "F", "D");
            SBO_VID_Currency.Properties.Settings.Default.Reload();
            // Set company properties

            switch (SBO_VID_Currency.Properties.Settings.Default.SQLType)    // 0->2005 1-2008 2-2012
            {
                case 0: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
                    break;
                case 1: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                    break;
                case 2: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                    break;
                case 3: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                    break;
                case 4: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                    break;
                case 5: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
                    break;
                case 6: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                    break;
                case 7: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                    break;
            }
            oCompany.UseTrusted = false;
            oCompany.Server = SBO_VID_Currency.Properties.Settings.Default.SBOServer; // +":40000";
            oCompany.DbUserName = SBO_VID_Currency.Properties.Settings.Default.DBUserName;
            oCompany.DbPassword = Crypt.Decrypt(SBO_VID_Currency.Properties.Settings.Default.DBPassword, "VIDCurrency0-09");
            oCompany.CompanyDB = CompanyName;
            oCompany.UserName = globals.SBOUser;
            oCompany.Password = globals.SBOPass;

            //Try to connect
            nErr = oCompany.Connect();
            oLog.LogMsg("try connect", "F", "D");

            if (nErr != 0) // if the connection failed
                oCompany.GetLastError(out nErr, out sErr);
            else
            {
                SAPbobsCOM.Recordset oRecordSet; // A recordset object
                oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    oRecordSet.DoQuery("select \"InstallNo\" from \"CINF\"");
                else
                    oRecordSet.DoQuery("select InstallNo from cinf");
                oLog.LogMsg("Select InstallNo", "F", "D");
                if (oRecordSet.EoF)
                {
                    sErr = "Installation Number no encontrado";
                    nErr = -1;
                    return;
                }

                string InstallId = (System.String)oRecordSet.Fields.Item("installNo").Value;

                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    oRecordSet.DoQuery("Select CURRENT_DATE \"Fec\" from DUMMY");
                else
                    oRecordSet.DoQuery("Select GetDate() Fec");
                DateTime Fec = (System.DateTime)(oRecordSet.Fields.Item("Fec").Value);
                TGlobalAddOnOptions FOpcionAddOn = new TGlobalAddOnOptions();
                List<String> FAddonsList = new List<String>();

                string oK = "C9A80FB4-D8C3-11E0-AEAD-7E944824019B";
                oLog.LogMsg("Before accept", "F", "D");

                Core = new VisualD.Core.TVisualDCore(InstallId, Fec, ref FOpcionAddOn, ref oK, FAddonsList);

                if (Core.ExitApp)
                {
                    sErr = "Licencia " + Core.Error;
                    nErr = -2;
                    oCompany.Disconnect();
                }
                else if (oK != "E6DCA176-D8C3-11E0-93B0-86944824019B")
                {
                    sErr = "VisualDCore enlace 2 inválido.";
                    nErr = -3;
                    oCompany.Disconnect();
                }
                else
                    sErr = "OK";

                if (Core.EsDEmo)
                    oLog.LogMsg("Licencia de demostracion " + Core.AddOnName, "F", "W");


                oLog.LogMsg("after accept", "F", "D");
            }

            oLog.LogMsg("Conexion a SBO: " + CompanyName + " - " + sErr, "F", "I");
        }

        private void processCurrency(SAPbobsCOM.Company oCompany)
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            String oSql;


             //Precondiciones iniciales (Moneda Local, Cambio Directo o Indirecto, decimales a redondear)

            if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                oSql = @"SELECT ""MainCurncy"", ""DirectRate"" , ""RateDec""   
                        FROM ""OADM""  ";
            else
                oSql = "SELECT MainCurncy, DirectRate , RateDec  " +
                       "FROM OADM  ";
            oRS.DoQuery(oSql);
            if (oRS.RecordCount > 0)
            {
                string mainCurrency = ((System.String)oRS.Fields.Item("MainCurncy").Value).Trim();
                string tipoCambio = ((System.String)oRS.Fields.Item("DirectRate").Value).Trim();
                int redondeo = ((System.Int32)oRS.Fields.Item("RateDec").Value);

                if (mainCurrency == "CLP" || mainCurrency == "$")  //si moneda sistema es pesos
                    deployCurrency(oCompany);
                else 
                {
                    if (tipoCambio == "N")  //si tipo de cambio es indirecto 
                        deployCurrencyCal(oCompany, mainCurrency);
                }
            }
            else
            {
                oLog.LogMsg("Parametros iniciales sin datos Tabla: OADM ", "A", "I");
            }





        }


        /// <summary>
        /// carga los valores calculados -> Moneda Definida Indirecta/ Monedas definidas en SAP
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="monedaIndirecta"> Moneda definida en SAP como indirecta</param>
        private void deployCurrencyCal(SAPbobsCOM.Company oCompany, string monedaIndirecta)
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.SBObob pSBObob = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            TasaCambioSBO oTasaCambioSBO = new TasaCambioSBO();
            List<TasaCambioSBO> oTasasCambio = new List<TasaCambioSBO>();
            DateTime Fecha = DateTime.Now;
            Dictionary<DateTime, Decimal> AuxDict = new Dictionary<DateTime, decimal>();

            String oSql;
            int diasAProcesar;
            double valorIndirecto = 0;

            try
            {


                // Validar definicion de monedas
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    oSql = @"SELECT ""CurrCode"", ""CurrName"" , ""ISOCurrCod"" , ""DocCurrCod"" , ""FrgnName""  
                        FROM ""OCRN"" 
                        WHERE ""Locked"" = 'N' ";
                else
                    oSql = "SELECT CurrCode, CurrName , ISOCurrCod , DocCurrCod , FrgnName " +
                           "FROM OCRN            " +
                           "WHERE Locked = 'N' ";

                oRS.DoQuery(oSql);
                if (oRS.RecordCount > 0)
                {

                    if (SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar > 0)
                        diasAProcesar = 0;
                    else
                        diasAProcesar = SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar;

                    oLog.LogMsg("Dias a procesar : " + diasAProcesar, "F", "D");

                    for (int x = diasAProcesar; x <= 1; x++)
                    {
                        Fecha = Fecha.AddDays(x);
                        string dateFormat = Fecha.ToString("yyyy-MM-dd");
                        oLog.LogMsg("Dia: " + dateFormat, "F", "D");
                        string url = SBO_VID_Currency.Properties.Settings.Default.WEBPage;
                        string responseRest = restGETWhitParameter(url, dateFormat);
                        if (responseRest != "[]")
                        {
                            var listTC = JsonConvert.DeserializeObject<List<TC>>(responseRest);

                            for (int y = 0; y < listTC.Count; y++)
                            {
                                if (monedaIndirecta == listTC[y].codigo)
                                    valorIndirecto = (Double)listTC[y].valor;
                            }
                            
                            for (int i = 0; i < oRS.RecordCount; i++)  //Monedas en SAP
                            {
                                string moneda = ((System.String)oRS.Fields.Item("CurrCode").Value).Trim();
                                string monedaISO = ((System.String)oRS.Fields.Item("ISOCurrCod").Value).Trim();
                                string monedaInternacional = ((System.String)oRS.Fields.Item("DocCurrCod").Value).Trim();
                                string calculoMoneda = ((System.String)oRS.Fields.Item("FrgnName").Value).Trim();

                                for (int y = 0; y < listTC.Count; y++) //monedas WS
                                {

                                    if (moneda == "CLP")  //asumo que siempre existirá CLP cuando se llama este metodo
                                    {
                                        oLog.LogMsg("procesado por CurrCode: " + Fecha + " Currency: " + moneda, "F", "D");
                                        UpdateSBOSAP(ref oCompany, Fecha, valorIndirecto, ref pSBObob, moneda);
                                        y = listTC.Count-1;
                                    }


                                    if (moneda == listTC[y].codigo)
                                    {
                                        if (calculoMoneda == "CLP")
                                        {
                                            oLog.LogMsg("procesado por CurrCode: " + Fecha + " Currency: " + moneda, "F", "D");
                                            UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                        }
                                        else
                                        {
                                            double val = valorIndirecto / (Double)listTC[y].valor;
                                            oLog.LogMsg("procesado por CurrCode Calculado: " + Fecha + " Currency: " + moneda, "F", "D");
                                            UpdateSBOSAP(ref oCompany, Fecha, val, ref pSBObob, moneda);
                                        }
                                        y = listTC.Count;
                                    }
                                    else
                                    {
                                        if (monedaISO == listTC[y].codigo)
                                        {
                                            if (calculoMoneda == "CLP")
                                            {
                                                oLog.LogMsg("procesado por ISOCurrCod: " + Fecha + " Currency: " + monedaISO, "F", "D");
                                                UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                            }
                                            else
                                            {
                                                double val = valorIndirecto / (Double)listTC[y].valor;
                                                oLog.LogMsg("procesado por ISOCurrCod Calculado: " + Fecha + " Currency: " + moneda, "F", "D");
                                                UpdateSBOSAP(ref oCompany, Fecha, val, ref pSBObob, moneda);
                                            }
                                            y = listTC.Count;
                                        }
                                        else
                                        {
                                            if (monedaInternacional == listTC[y].codigo)
                                            {
                                                if (calculoMoneda == "CLP")
                                                {
                                                    oLog.LogMsg("procesado por DocCurrCod: " + Fecha + " Currency: " + monedaInternacional, "F", "D");
                                                    UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                                }
                                                else
                                                {
                                                    double val = valorIndirecto / (Double)listTC[y].valor;
                                                    oLog.LogMsg("procesado por DocCurrCod Calculado: " + Fecha + " Currency: " + moneda, "F", "D");
                                                    UpdateSBOSAP(ref oCompany, Fecha, val, ref pSBObob, moneda);
                                                }
                                                y = listTC.Count;
                                            }
                                        }
                                    }
                                }
                                oRS.MoveNext();
                            }
                            oRS.MoveFirst();
                        }
                        Fecha = DateTime.Now;
                    }
                }
            }
            catch (Exception e)
            {
                oLog.LogMsg("deployCurrencyCal Error :  exepcion: " + e.Message, "A", "E");
            }

        }

        /// <summary>
        /// carga los valores que se encuentra en el Webservice
        /// </summary>
        /// <param name="oCompany"></param>
        private void deployCurrency(SAPbobsCOM.Company oCompany)
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.SBObob pSBObob = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            TasaCambioSBO oTasaCambioSBO = new TasaCambioSBO();
            List<TasaCambioSBO> oTasasCambio = new List<TasaCambioSBO>();
            DateTime Fecha = DateTime.Now;
            Dictionary<DateTime, Decimal> AuxDict = new Dictionary<DateTime, decimal>();

            String oSql;
            int diasAProcesar;
            
            // Validar definicion de monedas
            if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                oSql = @"SELECT ""CurrCode"", ""CurrName"" , ""ISOCurrCod"" , ""DocCurrCod""  
                        FROM ""OCRN"" 
                        WHERE ""Locked"" = 'N' ";
            else
                oSql = "SELECT CurrCode, CurrName , ISOCurrCod , DocCurrCod " +
                       "FROM OCRN            " +
                       "WHERE Locked = 'N' ";

            oRS.DoQuery(oSql);
            if (oRS.RecordCount > 0)
            {

                if (SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar > 0)
                    diasAProcesar = 0;
                else
                    diasAProcesar = SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar;

                oLog.LogMsg("Dias a procesar : " + diasAProcesar, "F", "D");

                for (int x = diasAProcesar; x <= 1; x++)
                {
                    Fecha = Fecha.AddDays(x);
                    string dateFormat = Fecha.ToString("yyyy-MM-dd");
                    oLog.LogMsg("Dia: " + dateFormat, "F", "D");
                    string url = SBO_VID_Currency.Properties.Settings.Default.WEBPage;
                    string responseRest = restGETWhitParameter(url, dateFormat);
                    if (responseRest != "[]")
                    {
                        var listTC = JsonConvert.DeserializeObject<List<TC>>(responseRest);
                        for (int i = 0; i < oRS.RecordCount; i++)  //Monedas en SAP
                        {
                            string moneda = ((System.String)oRS.Fields.Item("CurrCode").Value).Trim();
                            string monedaISO = ((System.String)oRS.Fields.Item("ISOCurrCod").Value).Trim();
                            string monedaInternacional = ((System.String)oRS.Fields.Item("DocCurrCod").Value).Trim();
                            for (int y = 0; y < listTC.Count; y++)
                            {
                                if (moneda == listTC[y].codigo)
                                {

                                    oLog.LogMsg("procesado por CurrCode: " + Fecha + " Currency: " + moneda, "F", "D");
                                    UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                    y = listTC.Count;
                                }
                                else
                                {
                                    if (monedaISO == listTC[y].codigo)
                                    {
                                        oLog.LogMsg("procesado por ISOCurrCod: " + Fecha + " Currency: " + monedaISO, "F", "D");
                                        UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                        y = listTC.Count;
                                    }
                                    else
                                    {
                                        if (monedaInternacional == listTC[y].codigo)
                                        {
                                            oLog.LogMsg("procesado por DocCurrCod: " + Fecha + " Currency: " + monedaInternacional, "F", "D");
                                            UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                            y = listTC.Count;
                                        }
                                    }
                                }
                            }
                            oRS.MoveNext();
                        }
                        oRS.MoveFirst();
                    }
                    Fecha = DateTime.Now;
                }
            }
        }

        private void processCurrencyDB(SqlConnection cnn)
        {
            TasaCambioSBO oTasaCambioSBO = new TasaCambioSBO();
            List<TasaCambioSBO> oTasasCambio = new List<TasaCambioSBO>();
            DateTime Fecha = DateTime.Now;
            Dictionary<DateTime, Decimal> AuxDict = new Dictionary<DateTime, decimal>();
            //String ApiKey = "c5656cd39657cf74083e0da48b1960e7963b4340";
            String oSql;
            int diasAProcesar;
            // Validar definicion de monedas
            oSql = "SELECT CurrCode, CurrName , ISOCurrCod   " +
                       "FROM OCRN            " +
                       "WHERE Locked = 'N' ";

            var command = new SqlCommand(oSql, cnn);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                if (reader.Read())
                {
                    DataTable dataTable = new DataTable();
                    dataTable.Load(reader);
                    
                if (SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar > 0)
                    diasAProcesar = 0;
                else
                    diasAProcesar = SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar;
        
                oLog.LogMsg("Dias a procesar : " + diasAProcesar , "F", "D");    
                    for (int x = diasAProcesar; x <= 1; x++)
                    {
                        Fecha = Fecha.AddDays(x);
                        string dateFormat = Fecha.ToString("yyyy-MM-dd");
                        oLog.LogMsg("Dia: " + dateFormat, "F", "D");
                        string url = SBO_VID_Currency.Properties.Settings.Default.WEBPage;
                        string responseRest = restGETWhitParameter(url, dateFormat);
                        if (responseRest != "[]")
                        {
                            var listTC = JsonConvert.DeserializeObject<List<TC>>(responseRest);
                            foreach (DataRow drow in dataTable.Rows)  //monedas en SAP
                            {
                                string moneda = System.Convert.ToString(drow["CurrCode"]);
                                string monedaISO = System.Convert.ToString(drow["ISOCurrCod"]);
                                for (int y = 0; y < listTC.Count; y++)
                                {
                                    if (moneda == listTC[y].codigo)
                                    {
                                        oLog.LogMsg("procesar por CurrCode: " + Fecha + " Currency: " + moneda, "F", "D");
                                        UpdateSBOSAPBD(ref cnn, Fecha, (Double)listTC[y].valor, moneda);
                                        y = listTC.Count;
                                    }
                                    else
                                    {
                                        if (monedaISO == listTC[y].codigo)
                                        {
                                            oLog.LogMsg("procesar por ISOCurrCod: " + Fecha + " Currency: " + moneda, "F", "D");
                                            UpdateSBOSAPBD(ref cnn, Fecha, (Double)listTC[y].valor, moneda);
                                            y = listTC.Count;
                                        }
                                    }
                                }
                            }
                        }
                        Fecha = DateTime.Now;
                    }
                }
            }
        }

        private void UpdateSBOSAPBD(ref SqlConnection cnn, DateTime FecActSBO, Double USDObs, string moneda)
        {
            int count = 0;
            try
            {
                string dateFormat = FecActSBO.ToString("yyyyMMdd");
                string s =  "SELECT count(*) FROM ORTT WHERE Currency = '{0}' AND RateDate = '{1}' ";
                s = String.Format(s, moneda, dateFormat);
                using (SqlCommand command = new SqlCommand(s, cnn))
                {
                    count = (int)command.ExecuteScalar();
                    if (count == 0) //si no existe tpo cambio para ese dia 
                    {
                        s = "INSERT INTO ORTT (RateDate,Currency,Rate,DataSource,UserSign) VALUES ('{0}','{1}',{2},'{3}',{4})";
                        s = String.Format(s, dateFormat, moneda, USDObs, 'P', 1);
                        command.CommandText = s;
                        try
                        {
                            command.ExecuteNonQuery();
                            oLog.LogMsg("Valor cargado Moneda: " + moneda + " Fecha:" + FecActSBO + " Valor: " + USDObs, "A", "I");
                        }
                        catch (Exception e)
                        {
                            oLog.LogMsg("Error al actualizar tasa de cambio en SBO SCRIPT " + e.Message, "A", "E");
                        }
                    }
                    else 
                    {
                        oLog.LogMsg("Tipo de cambio  " + moneda + " ya se encuentra ingresado en la sociedad ", "A", "D");
                    }
                }
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error al actualizar tasa de cambio en SBO - Moneda: " + moneda + " exepcion: "  + e.Message, "A", "E");
            }
        }

        private void UpdateSBOSAP(ref SAPbobsCOM.Company oCompany, DateTime FecActSBO, Double USDObs, ref SAPbobsCOM.SBObob oSBObob, string moneda)
        {

            try
            {
                //if (SBO_VID_Currency.Properties.Settings.Default.TipoCambio == 1)
                //{
                //    cambioIndirecto = true;
                //    if (SBO_VID_Currency.Properties.Settings.Default.MonedaBaseIndirecta == "E")
                //        MonedaBase = EU;
                //    else
                //        MonedaBase = USDObs;
                //}
                //string s;

                try
                {
                    //s = SBO_VID_Currency.Properties.Settings.Default.Dolar;
                    if ((moneda != "") && (moneda != null))
                    {
                        oSBObob.SetCurrencyRate(moneda, FecActSBO, USDObs, false);
                        oLog.LogMsg("Valor a cargar Moneda: " + moneda + " Fecha:" + FecActSBO + " Valor: " + USDObs, "A", "I");
                    }
                }
                catch (Exception e)
                {
                    oLog.LogMsg("Tipo de cambio  " +  moneda + " ya se encuentra ingresado en la sociedad " +  e.Message, "A", "D");
                }


            }
            catch (Exception e)
            {
                oLog.LogMsg("Error al actualizar tasa de cambio en SBO - " + e.Message, "A", "E");
            }
        }


        private void UpdateSBO(ref SAPbobsCOM.Company oCompany, DateTime FecActSBO, Double USDObs, Double EU, Double Uefe, ref SAPbobsCOM.SBObob oSBObob)
        {
            Double MonedaBase = 0; ;
            bool cambioIndirecto = false;

            try
            {
                if (SBO_VID_Currency.Properties.Settings.Default.TipoCambio == 1)
                {
                    cambioIndirecto = true;
                    if (SBO_VID_Currency.Properties.Settings.Default.MonedaBaseIndirecta == "E")
                        MonedaBase = EU;
                    else
                        MonedaBase = USDObs;
                }

                string s;

                try
                {
                    s = SBO_VID_Currency.Properties.Settings.Default.Dolar;
                    if ((USDObs >= 0) && (s != "") && (s != null))
                        oSBObob.SetCurrencyRate(s, FecActSBO, USDObs, false);
                }
                catch
                {
                }
                try
                {
                    s = SBO_VID_Currency.Properties.Settings.Default.Euro;
                    if ((EU >= 0) && (s != "") && (s != null))
                        oSBObob.SetCurrencyRate(s, FecActSBO, EU, false);
                }
                catch
                {
                }
                try
                {
                    s = SBO_VID_Currency.Properties.Settings.Default.UF;
                    if ((Uefe >= 0) && (s != "") && (s != null))
                        oSBObob.SetCurrencyRate(s, FecActSBO, Uefe, false);
                }
                catch
                {
                }
                try
                {
                    s = SBO_VID_Currency.Properties.Settings.Default.DolarI;
                    if ((USDObs > 0) && (s != "") && (s != null) && (MonedaBase > 0) && (cambioIndirecto))
                        oSBObob.SetCurrencyRate(s, FecActSBO, MonedaBase / USDObs, false);
                }
                catch
                {
                }
                try
                {
                    s = SBO_VID_Currency.Properties.Settings.Default.EuroI;
                    if ((EU > 0) && (s != "") && (s != null) && (MonedaBase > 0) && (cambioIndirecto))
                        oSBObob.SetCurrencyRate(s, FecActSBO, MonedaBase / EU, false);
                }
                catch
                {
                }
                try
                {
                    s = SBO_VID_Currency.Properties.Settings.Default.UFI;
                    if ((Uefe > 0) && (s != "") && (s != null) && (MonedaBase > 0) && (cambioIndirecto))
                        oSBObob.SetCurrencyRate(s, FecActSBO, MonedaBase / Uefe, false);
                }
                catch
                {
                }
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error al actualizar tasa de cambio en SBO - " + e.Message, "A", "E");
            }
        }
    }
}
