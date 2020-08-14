using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SBO_VID_Currency
{
    public partial class SetupParams : Form
    {
        bool ConexionOK = false;
        Color OrigBackColor;
        Color OrigForeColor;

        public SetupParams()
        {
            InitializeComponent();
            this.DialogResult = System.Windows.Forms.DialogResult.None;
            OrigBackColor = statusStrip1.BackColor;
            OrigForeColor = statusStrip1.ForeColor;
            readSQLServer();
        }

        private void readSQLServer()
        {
            btnOk.Text = "Conectar";
            cbSqlType.SelectedIndex = SBO_VID_Currency.Properties.Settings.Default.SQLType; // 0->2005 1-2008 2-2012
            tbServidor.Text = SBO_VID_Currency.Properties.Settings.Default.SBOServer;
            tbDBUser.Text = SBO_VID_Currency.Properties.Settings.Default.DBUserName;
            if (SBO_VID_Currency.Properties.Settings.Default.DBPassword.Length != 24)
                tbDBPass.Text = "";
            else
                tbDBPass.Text = Crypt.Decrypt(SBO_VID_Currency.Properties.Settings.Default.DBPassword, "VIDCurrency0-09");
            toolStripStatusLabel1.Text = "Conectar a servidor";
        }

        private bool GetCompanies()
        {
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company(); // The company object
            SAPbobsCOM.Recordset oRecordSet;

            statusStrip1.BackColor = Color.LightSalmon;
            statusStrip1.ForeColor = Color.Black;
            toolStripStatusLabel1.Text = "Buscando servidor";

            switch (cbSqlType.SelectedIndex)
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
                case 5: oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                    break;
            }
            oCompany.UseTrusted = false;
            oCompany.Server = tbServidor.Text;
            oCompany.DbUserName = tbDBUser.Text;
            oCompany.DbPassword = tbDBPass.Text;

            try
            {
                int nErr;
                string sErr;

                oRecordSet = oCompany.GetCompanyList();
                oCompany.GetLastError(out nErr, out sErr);
                if (nErr != 0)
                {
                    statusStrip1.BackColor = Color.Orange;
                    statusStrip1.ForeColor = Color.White;
                    toolStripStatusLabel1.Text = "Error en conexión al servidor";
                    MessageBox.Show(sErr);
                    return false;
                }
                else
                {
                    statusStrip1.BackColor = OrigBackColor;
                    statusStrip1.ForeColor = OrigForeColor;
                    toolStripStatusLabel1.Text = "Conectado al servidor";
                    FillAndCompareCompanies(oRecordSet);
                    return true;
                }
            }
            catch (Exception e)
            {
                statusStrip1.BackColor = Color.Orange;
                statusStrip1.ForeColor = Color.White;
                toolStripStatusLabel1.Text = "Error en conexión al servidor";
                MessageBox.Show(e.Message);
                return false;
            }
        }

        private void FillAndCompareCompanies(SAPbobsCOM.Recordset oRecordSet)
        {
            cblCompanias.Items.Clear();
            while (!oRecordSet.EoF)
            {
                cblCompanias.Items.Add(oRecordSet.Fields.Item(0).Value);
                oRecordSet.MoveNext();
            }

            int i;
            string[] separador = new string[] { ";" };
            string[] lista = SBO_VID_Currency.Properties.Settings.Default.Companies.Split(separador, StringSplitOptions.RemoveEmptyEntries);
            foreach (String s in lista)
            {
                i = cblCompanias.Items.IndexOf(s);
                if (i > -1)
                    cblCompanias.SetItemChecked(i, true);
            }
        }

        private void readParams()
        {
            tbSBOUser.Text = SBO_VID_Currency.Properties.Settings.Default.SBOUserName;
            if (SBO_VID_Currency.Properties.Settings.Default.SBOPassword.Length != 24)
                tbSBOPass.Text = "";
            else
                tbSBOPass.Text = Crypt.Decrypt(SBO_VID_Currency.Properties.Settings.Default.SBOPassword, "VIDCurrency0-09");
            tbPaginaWEB.Text = SBO_VID_Currency.Properties.Settings.Default.WEBPage;
            tbReintento.Text = SBO_VID_Currency.Properties.Settings.Default.Reintento_Seg.ToString();
            tbDias.Text = SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar.ToString();
            cbConsola.Checked = SBO_VID_Currency.Properties.Settings.Default.Consola;
            cbScheduler.Checked = SBO_VID_Currency.Properties.Settings.Default.UseScheduler;
            cbTipoCambio.SelectedIndex = SBO_VID_Currency.Properties.Settings.Default.TipoCambio;
            tbDDolar.Text = SBO_VID_Currency.Properties.Settings.Default.Dolar;
            tbDEuro.Text = SBO_VID_Currency.Properties.Settings.Default.Euro;
            tbDUF.Text = SBO_VID_Currency.Properties.Settings.Default.UF;

            if (cbTipoCambio.SelectedIndex == 1)
            {
                panel1.Enabled = true;
                tbIDolar.Text = SBO_VID_Currency.Properties.Settings.Default.DolarI;
                tbIEuro.Text = SBO_VID_Currency.Properties.Settings.Default.EuroI;
                tbIUF.Text = SBO_VID_Currency.Properties.Settings.Default.UFI;
                switch (SBO_VID_Currency.Properties.Settings.Default.MonedaBaseIndirecta)
                {
                    case "E": radioButton2.Checked = true;
                        break;
                    case "U": radioButton3.Checked = true;
                        break;
                    default : radioButton1.Checked = true;
                        break;
                }
            }
            else
            {
                panel1.Enabled = false;
                tbIDolar.Text = "";
                tbIEuro.Text = "";
                tbIUF.Text = "";
            }
            groupBox2.Enabled = true;
            groupBox3.Enabled = true;
            tbReintento.Enabled = true;
            tbDias.Enabled = true;
            cbConsola.Enabled = false;
            cbScheduler.Enabled = false;
        }

        private void SaveParams()
        {
            SBO_VID_Currency.Properties.Settings.Default.SQLType = cbSqlType.SelectedIndex;
            SBO_VID_Currency.Properties.Settings.Default.SBOServer = tbServidor.Text;
            SBO_VID_Currency.Properties.Settings.Default.DBUserName = tbDBUser.Text;
            SBO_VID_Currency.Properties.Settings.Default.DBPassword = Crypt.Encrypt(tbDBPass.Text, "VIDCurrency0-09");

            SBO_VID_Currency.Properties.Settings.Default.SBOUserName = tbSBOUser.Text;
            SBO_VID_Currency.Properties.Settings.Default.SBOPassword = Crypt.Encrypt(tbSBOPass.Text, "VIDCurrency0-09");
            // SBO_VID_Currency.Properties.Settings.Default.WEBPage = tbPaginaWEB.Text;
            SBO_VID_Currency.Properties.Settings.Default.Reintento_Seg = Convert.ToInt32(tbReintento.Text, 10);
            SBO_VID_Currency.Properties.Settings.Default.DiasAnterioresAProcesar = Convert.ToInt32(tbDias.Text, 10);
            // SBO_VID_Currency.Properties.Settings.Default.Consola = cbConsola.Checked;
            // SBO_VID_Currency.Properties.Settings.Default.UseScheduler = cbScheduler.Checked;
            SBO_VID_Currency.Properties.Settings.Default.TipoCambio = cbTipoCambio.SelectedIndex;
            SBO_VID_Currency.Properties.Settings.Default.Dolar = tbDDolar.Text;
            SBO_VID_Currency.Properties.Settings.Default.Euro = tbDEuro.Text;
            SBO_VID_Currency.Properties.Settings.Default.UF = tbDUF.Text;
            if (cbTipoCambio.SelectedIndex == 1)
            {
                SBO_VID_Currency.Properties.Settings.Default.DolarI = tbIDolar.Text;
                SBO_VID_Currency.Properties.Settings.Default.EuroI = tbIEuro.Text;
                SBO_VID_Currency.Properties.Settings.Default.UFI = tbIUF.Text;
            }
            else
            {
                SBO_VID_Currency.Properties.Settings.Default.DolarI = "";
                SBO_VID_Currency.Properties.Settings.Default.EuroI = "";
                SBO_VID_Currency.Properties.Settings.Default.UFI = "";
            }

            if (radioButton2.Checked)
                SBO_VID_Currency.Properties.Settings.Default.MonedaBaseIndirecta = "E";
            else if (radioButton3.Checked)
                SBO_VID_Currency.Properties.Settings.Default.MonedaBaseIndirecta = "U";
            else 
                SBO_VID_Currency.Properties.Settings.Default.MonedaBaseIndirecta = "D";

            SBO_VID_Currency.Properties.Settings.Default.Companies = "";
            bool first = true;
            for (int i = 0; i <= cblCompanias.Items.Count - 1; i++)
            {
                if (cblCompanias.GetItemChecked(i))
                {
                    if (first)
                    {
                        SBO_VID_Currency.Properties.Settings.Default.Companies = (string)cblCompanias.Items[i];
                        first = false;
                    }
                    else
                        SBO_VID_Currency.Properties.Settings.Default.Companies = SBO_VID_Currency.Properties.Settings.Default.Companies + ";" + (string)cblCompanias.Items[i];
                }
            }
            SBO_VID_Currency.Properties.Settings.Default.Save();
        }

        private void cbTipoCambio_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbTipoCambio.SelectedIndex == 1)
                panel1.Enabled = true;
            else
            {
                panel1.Enabled = false;
                tbIDolar.Text = "";
                tbIEuro.Text = "";
                tbIUF.Text = "";
            }

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!ConexionOK)
            {
                if (!GetCompanies())
                {
                    this.DialogResult = System.Windows.Forms.DialogResult.None;
                    return;
                }
                ConexionOK = true;
                groupBox1.Enabled = false;
                btnOk.Text = "Ok";
                readParams();
                this.DialogResult = System.Windows.Forms.DialogResult.None;
                return;
            }
            SaveParams();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            //            printForm1.Form = this;
            //            printForm1.PrintAction = System.Drawing.Printing.PrintAction.PrintToPrinter;
            //            printForm1.Print();
        }

        private void SetupParams_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult == System.Windows.Forms.DialogResult.None)
                e.Cancel = true;
            else
                e.Cancel = false;
        }
    }
}
