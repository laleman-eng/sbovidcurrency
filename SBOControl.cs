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
        private void GetParameters(SAPbobsCOM.Company oCompany)
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
            int i;

            try
            {
                if (oCompany == null)
                    oCompany = new SAPbobsCOM.Company();

                GetParameters(oCompany);

                for (i = 0; i < globals.CompanyList.Count; i++)
                {
                    if (!oCompany.Connected)
                    {
                        ConnectSBO(ref oCompany, ref nErr, ref sErr, (String)globals.CompanyList[i]);
                        if (nErr != 0)
                        {
                            oLog.LogMsg(sErr, "F", "E");
                            oCompany.Disconnect();
                            return;
                        }
                    }

                    if (oCompany.Connected)
                    {
                            processCurrency(oCompany);
                        oCompany.Disconnect();
                    }
                }
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error en ejecución de servicio " + e.Message, "F", "E");
                if (oCompany.Connected)
                    oCompany.Disconnect();
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
                case 6: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
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
                    oLog.LogMsg("Licencia de demostración " + Core.AddOnName, "F", "E");


                oLog.LogMsg("after accept", "F", "D");
            }

            oLog.LogMsg("Conexión a SBO: " + CompanyName + " - " + sErr, "F", "E");
        }

        private void processCurrency(SAPbobsCOM.Company oCompany)
        {
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.SBObob pSBObob = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            TasaCambioSBO oTasaCambioSBO = new TasaCambioSBO();
            List<TasaCambioSBO> oTasasCambio = new List<TasaCambioSBO>();
            DateTime Fecha = DateTime.Now;
            Int32 year;
            Int32 month;
            Int32 day;
            Decimal valor;

            Dictionary<DateTime, Decimal> AuxDict = new Dictionary<DateTime, decimal>();
            String ApiKey = "c5656cd39657cf74083e0da48b1960e7963b4340";
            String sFecha;
            String oSql;
            String Moneda;

            Boolean bDolar = false;
            Boolean bEuro = false;
            Boolean bUF = false;
            Boolean bDolarI = false;
            Boolean bEuroI = false;
            Boolean bUFI = false;

            // Validar definicion de monedas
            if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                oSql = @"SELECT ""CurrCode"", ""CurrName""    
                        FROM ""OCRN"" WHERE ""Locked"" = 'N' ";
            else
                oSql = "SELECT CurrCode, CurrName  " +
                       "FROM OCRN            " +
                       "WHERE Locked = 'N' ";

            oRS.DoQuery(oSql);
            if (oRS.RecordCount > 0)
            {
                int j = 1;
                if (SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar != 0)
                    j = 2;
                for (int x = 0; x < j; x++)
                {
                    if (x == 1)
                        Fecha = Fecha.AddDays(SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar);
                    
                    string dateFormat = Fecha.ToString("yyyy-MM-dd");
                    string url = SBO_VID_Currency.Properties.Settings.Default.WEBPage;
                    string responseRest = restGETWhitParameter(url, dateFormat);
                    if (responseRest != "")
                    {
                        var listTC = JsonConvert.DeserializeObject<List<TC>>(responseRest);

                        for (int i = 0; i < oRS.RecordCount; i++)  //Monedas en SAP
                        {
                            string moneda = ((System.String)oRS.Fields.Item("CurrCode").Value).Trim();

                            for (int y = 0; y < listTC.Count; y++)
                            {
                                if (moneda == listTC[y].codigo)
                                {
                                    UpdateSBOSAP(ref oCompany, Fecha, (Double)listTC[y].valor, ref pSBObob, moneda);
                                    y = listTC.Count;
                                }
                            }
                            oRS.MoveNext();
                        }
                        oRS.MoveFirst();
                    }
                }

                sFecha = Fecha.ToString("yyyyMMdd");
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    oRS.DoQuery("select distinct \"RateDate\" from  \"ORTT\" " +
                                " where \"RateDate\" >= TO_DATE('" + sFecha + "', 'YYYYMMDD') " +
                                " order by \"RateDate\" ");
                else
                    oRS.DoQuery("select distinct RateDate from  ORTT " +
                                " where RateDate >= CONVERT(datetime,'" + sFecha + "',112) " +
                                " order by RateDate ");
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
                string s;

                try
                {
                    //s = SBO_VID_Currency.Properties.Settings.Default.Dolar;
                    if  ((moneda != "") && (moneda != null))
                        oSBObob.SetCurrencyRate(moneda, FecActSBO, USDObs, false);
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
