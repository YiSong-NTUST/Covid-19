using SmartTagTool;
using System;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Windows.Forms;
//using Thorlabs.TLPM_64.Interop;
using Thorlabs.PM100D_64.Interop;
using System.Threading;
using System.Threading.Tasks;

namespace Spectrum_Test
{
    public delegate void AddLogDelegate(String myString); 

    public partial class MainForm : Form
    {
        //private TLPM tlpm;
        private PM100D pm100d;
        double powerValue;
        double sum = 0.0;
        bool Avg_flag = false;
        System.Timers.Timer tmrPowerMeter = new System.Timers.Timer();
        int timer_count = 0; //平均次數

        public AddLogDelegate logDelegate;
        public AddLogDelegate rawlogDelegate;

        string selPort;

        char[] recv_buff = new char[1024];
        int recv_count = 0;

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            if (serialPort.IsOpen) return;

            loadPorts();
        }

        private void loadPorts()
        {
            string[] ports = SerialPort.GetPortNames();

            cboPort.Items.Clear();
            foreach (string portName in ports)
            {
                bool found = false;
                foreach (Object obj in cboPort.Items)
                {
                    if (((string)obj).Equals(portName))
                    {
                        found = true;
                        break;
                    }

                }
                if (!found)
                {
                    cboPort.Items.Add(portName);
                    if (selPort != "" && selPort.Equals(portName))
                    {
                        cboPort.SelectedIndex = cboPort.Items.Count - 1;
                    }
                }
            }
        }


        private void _rawlog(string message)
        {
            txtLog.AppendText(message);

            txtLog.SelectionStart = txtLog.Text.Length - 1; // add some logic if length is 0
            txtLog.SelectionLength = 0;
        }

        private void _log(string message)
        {
            _rawlog("\r\n" + message + "\r\n");
        }

        private void Log(string message)
        {
            txtLog.Invoke(this.logDelegate, new Object[] { message });
        }

        private void RawLog(string message)
        {
            txtLog.Invoke(this.rawlogDelegate, new Object[] { message });
        }

        private IniFile GetIniFile()
        {
            string iniPath = string.Format("{0}\\setting.ini", Application.StartupPath);

            if (!File.Exists(iniPath))
            {
                StreamWriter sw = File.CreateText(iniPath);
                sw.Close();
            }


            IniFile ini = new IniFile(iniPath);

            return ini;
        }

        private void LoadSetting()
        {
            IniFile ini = GetIniFile();

            // load setting
            selPort = ini.IniReadValue("SETTING", "COMPORT");
        }

        private void SaveSetting()
        {
            // load setting
            if (cboPort.SelectedItem != null)
                selPort = cboPort.SelectedItem.ToString();

            /*test_led1 = chk_test_led1.Checked;
            test_led2 = chk_test_led2.Checked;
            test_led3 = chk_test_led3.Checked;

            test_led_sec = int.Parse(txt_test_led_time.Text);*/


            IniFile ini = GetIniFile();

            // load setting
            ini.IniWriteValue("SETTING", "COMPORT", selPort);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {    
            selPort = "";
            //this.logDelegate = new AddLogDelegate(_log);
            //this.rawlogDelegate = new AddLogDelegate(_rawlog);

            //LoadSetting();


            //loadPorts();
            initPowerMeter();
            
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            string selPort = (string)cboPort.SelectedItem;

            if (selPort == null || selPort.Equals("")) return;

            serialPort.PortName = selPort;
            //serialPort.BaudRate = 460800;
            serialPort.BaudRate = 115200;
            serialPort.DataBits = 8;
            serialPort.Handshake = Handshake.None;
            serialPort.StopBits = StopBits.One;

            try
            {
                serialPort.Open();

                serialPort.DataReceived += SerialPort_DataReceived;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information");
                return;
            }

  
            btnOpen.Visible = false;
            btnClose.Visible = true;

        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            byte data;
            SerialPort sp = (SerialPort)sender;

            char[] array = new char[4096];
            int i = 0;

            while (sp.BytesToRead > 0)
            {
                data = (byte)sp.ReadByte();
                array[i++] = (char)data;

                if (recv_count > 0 || data == '$')
                    recv_buff[recv_count++] = (char)data;
            }

            array[i] = '\0';

            Log(new string(array));
            //this.Invoke(this.logDelegate, new string(array));            
        }

        private void initPowerMeter()
        {
            comboBox1.SelectedIndex = 0;
            pm100d = new PM100D("USB0::0x1313::0x8078::P0025621::INSTR", false, false);

            double Wavelength;
            pm100d.getWavelength(0,out Wavelength);

            for(int i = 0; i < comboBox1.Items.Count; i++)
            {
                string valCmb = comboBox1.Items[i].ToString();
                string val = valCmb.Remove(valCmb.Length - 3, 3);

                if (Convert.ToDouble(val) == Wavelength)
                {
                    comboBox1.SelectedIndex = i;
                }
            }

            tmrPowerMeter.Interval = 250;
            tmrPowerMeter.Elapsed += new System.Timers.ElapsedEventHandler(_TimersTimer_Elapsed);
            tmrPowerMeter.SynchronizingObject = this;
            tmrPowerMeter.Start();
        }

        void _TimersTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            
            try
            {
                /*HandleRef Instrument_Handle = new HandleRef();
                  TLPM searchDevice = new TLPM(Instrument_Handle.Handle);
                  tlpm = new TLPM("USB0::0x1313::0x8078::P0025621::0::INSTR", false, false);  //  For valid Ressource_Name see NI-Visa documentation.  
                  tlpm.measPower(out powerValue);*/
                               
                pm100d.measPower(out powerValue);
                lbPower.Text = (powerValue * 1000.0).ToString("#0.000");
                lbPower.ForeColor = System.Drawing.Color.Black;
            }
            catch (BadImageFormatException bie)
            {
                lbPower.Text = bie.Message;
            }
            catch (NullReferenceException nre)
            {
                lbPower.Text = nre.Message;
            }
            catch (ExternalException ex)
            {
                lbPower.Text = ex.Message;
            }
            finally
            {
                /*if (pm100d != null)
                    pm100d.Dispose();*/
            }

            if (Avg_flag)
            {
                int num = Convert.ToInt32(txtAvg.Text);
                sum += (powerValue * 1000.0);
                timer_count++;

                if (timer_count == num)
                {
                    tmrPowerMeter.Enabled = false;

                    sum /= Convert.ToDouble(num);
                    lbPower.InvokeIfRequired(() =>
                    {
                        lbPower.Text = sum.ToString("#0.000");
                        lbPower.ForeColor = System.Drawing.Color.Red;
                    });

                    sum = 0.0;
                    timer_count = 0;
                    btnAvg.Text = "開始";
                    btnAvg.Enabled = true;
                }
            }                
        }

        private void btnAvg_Click(object sender, EventArgs e)
        {
            if (Avg_flag)
            {
                Avg_flag = false;
                tmrPowerMeter.Enabled = true;
                btnAvg.Text = "停止";
            }
            else
            {
                Avg_flag = true;
                btnAvg.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string valCmb = comboBox1.SelectedItem.ToString();
            string val = valCmb.Remove(valCmb.Length - 3, 3);
            pm100d.setWavelength(Convert.ToDouble(val));
        }

        private void txtAvg_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
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
