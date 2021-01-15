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
    public partial class ServiceForm : Form
    {
        private Timer _timer = new Timer();
        private SBOControl SBOCtrl;
        private Boolean FirstTime;
        private DateTime LastUpdateDateCurrency;
        private int HourOfInitUpdate;
        public Logs.Logger oLog;

        public ServiceForm()
        {
            InitializeComponent();
            Start();
        }

        private void Start()
        {
            SBOCtrl = new SBOControl();
            oLog = new Logs.Logger();
            FirstTime = true;
            LastUpdateDateCurrency = new DateTime(2010, 1, 1);

            HourOfInitUpdate = SBO_VID_Currency.Properties.Settings.Default.HoraInicio;
            HourOfInitUpdate = (HourOfInitUpdate > 23) ? 23 : HourOfInitUpdate;
            HourOfInitUpdate = (HourOfInitUpdate < 18) ? 18 : HourOfInitUpdate;

            oLog.LogMsg("Servicio iniciado", "A", "I");


            SBOCtrl.oLog = oLog;
            _timer.Interval = 3 * 1000; // 5 Minutos = 300 Seg. - Primera vez espera 3 seg.
            _timer.Tick += OnTimerTick;

            if (SBO_VID_Currency.Properties.Settings.Default.Consola)
            {
                _timer.Start();
                oLog.TextBoxMsg = null;
                oLog.LogMsg("Timer enabled", "F", "I");
            }
            else
            {
                oLog.TextBoxMsg = this.tbLog;
            }
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            _timer.Stop();

            int nError = 0;
            string sMsg = "";

            if (!SBO_VID_Currency.Properties.Settings.Default.UseScheduler)
            {
                if (LastUpdateDateCurrency.Date > DateTime.Now.Date) 
                    return;
                if ((LastUpdateDateCurrency.Date == DateTime.Now.Date) && (DateTime.Now.Hour < HourOfInitUpdate))
                    return;
            }
            else
            {
                if (LastUpdateDateCurrency.Date > DateTime.Now.Date)
                    System.Windows.Forms.Application.Exit();
                if ((LastUpdateDateCurrency.Date == DateTime.Now.Date) && (DateTime.Now.Hour < HourOfInitUpdate))
                    System.Windows.Forms.Application.Exit();
            }
            try
            {
                oLog.LogMsg("Timer paused", "F", "D");

                if (FirstTime)
                {
                    FirstTime = false;
                    _timer.Interval = SBO_VID_Currency.Properties.Settings.Default.Reintento_Seg * 1000;
                }

                oLog.LogMsg("Before Doit", "F", "D");
                SBOCtrl.Doit(ref nError, ref sMsg);
                //System.Environment.Exit(0); //cierrpo app
                //SBOCtrl.Doit(ref nError, ref sMsg, ref LastUpdateDateCurrency); antes
            }
            catch (Exception eX)
            {
                oLog.LogMsg("Error Servicio : " + eX.Message, "A", "E");
            }
            //_timer.Enabled = true;
            //oLog.LogMsg("Timer restart", "F", "D");
        }

        private void btnInit_Click(object sender, EventArgs e)
        {
            btnFin.Enabled = true;
            btnInit.Enabled = false;
            _timer.Start();
            oLog.LogMsg("timer iniciado", "A", "I");
        }

        private void btnFin_Click(object sender, EventArgs e)
        {
            btnInit.Enabled = true;
            btnFin.Enabled = false;
            _timer.Stop();
            oLog.LogMsg("timer detenido", "A", "I");
        }

        private void ServiceForm_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            if (SBO_VID_Currency.Properties.Settings.Default.Consola)
                this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetupParams F1 = new SetupParams();
            var result = F1.ShowDialog(this);
            if (result == DialogResult.OK)
            {
            }
        }
    }
}
