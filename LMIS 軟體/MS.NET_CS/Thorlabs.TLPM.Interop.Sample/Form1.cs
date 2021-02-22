namespace Thorlabs.TLPM_32.Interop.Sample
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Forms;
    using Thorlabs.PM100D_64.Interop;

    public partial class Form1 : Form
    {
        private PM100D tlpm;

        System.Timers.Timer tmrPowerMeter = new System.Timers.Timer();

        public Form1()
        {
            InitializeComponent();

            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tmrPowerMeter.Interval = 200;
            tmrPowerMeter.Elapsed += new System.Timers.ElapsedEventHandler(_TimersTimer_Elapsed);
            //tmrPowerMeter.SynchronizingObject = this;
            tmrPowerMeter.Start();
        }

        void _TimersTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                HandleRef Instrument_Handle = new HandleRef();

                PM100D searchDevice = new PM100D(Instrument_Handle.Handle);

                uint count = 0;
                string firstPowermeterFound = "";

                System.Diagnostics.Debug.WriteLine("click debug!!");
                try
                {
                    int pInvokeResult = searchDevice.findRsrc(out count);

                    if (count > 0)
                    {
                        StringBuilder descr = new StringBuilder(1024);
                        searchDevice.getRsrcName(0, descr);
                        firstPowermeterFound = descr.ToString();
                    }
                }
                catch { }

                if (count == 0)
                {
                    searchDevice.Dispose();
                    //lbPower.Text = "No power meter could be found.";
                    return;
                }
                tlpm = new PM100D(firstPowermeterFound, false, false);  //  For valid Ressource_Name see NI-Visa documentation.
                double powerValue;
                int err = tlpm.measPower(out powerValue);

                labelPower.InvokeIfRequired(() =>
                {
                    labelPower.Text = (powerValue * 1000.0).ToString("0.0000");
                });
            }
            catch (BadImageFormatException bie)
            {
                labelPower.Text = bie.Message;
            }
            catch (NullReferenceException nre)
            {
                labelPower.Text = nre.Message;
            }
            catch (ExternalException ex)
            {
                labelPower.Text = ex.Message;
            }
            finally
            {
                if (tlpm != null)
                    tlpm.Dispose();
            }
         }
    }

    //擴充方法
    public static class Extension
    {
        //非同步委派更新UI
        public static void InvokeIfRequired(this Control control, MethodInvoker action)
        {
            if (control.InvokeRequired)//在非當前執行緒內 使用委派
            {
                control.Invoke(action);
            }
            else
            {
                action();
            }
        }
    }
}
