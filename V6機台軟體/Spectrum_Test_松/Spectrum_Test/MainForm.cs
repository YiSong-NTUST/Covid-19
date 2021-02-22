using SmartTagTool;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Spectrum_Test
{
    public delegate void AddLogDelegate(String myString); 

    

    public partial class MainForm : Form
    {
        Stopwatch sw = new Stopwatch();
        public static string command = "";
        public static List<List<int>> ALL_POINT_CAL = new List<List<int>>();
        public static List<byte> CAL = new List<byte>();
        public static List<int> MULT_CAL = new List<int>();

        public const int CMD_RET_OK = 0;
        public const int CMD_RET_TIMEOUT = 1;
        public const int CMD_RET_NACK = 2;
        public const int CMD_RET_POS_ON = 3;
        public const int CMD_RET_POS_OFF = 4;
        public const int CMD_RET_WDIR = 5;

        public const int CMD_RET_ERR = 255;

        public AddLogDelegate logDelegate;
        public AddLogDelegate rawlogDelegate;

        string selPort;
        double T1, T2, XSC, XTS, XAS, X1, total_points, step;
        int DG, AG, EXP_max, EXP_init, I_max, I_thr;
        int maxValue, EXP, EXP1, I1, EXP2, I2;
        int CAL_cycle, wl;

        char[] recv_buff = new char[24567];
        int recv_count = 0;
        int cycle = 1;

        bool flag = false;

        private int iTask = 0;
        int ret;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            selPort = "";
            this.logDelegate = new AddLogDelegate(_log);
            this.rawlogDelegate = new AddLogDelegate(_rawlog);

            chart1.ChartAreas[0].AxisY.Maximum = 4095;

            LoadSetting();
            loadPorts();
        }

        /**  port 更新   */
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            if (serialPort.IsOpen) return;

            loadPorts();
        }

        /** 載入port */
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

        /** 開始進行QC程序 */
        private async void startbtn_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                while (true)
                {
                    switch (iTask)
                    {
                        /** 設方向 */
                        case 0:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "設方向";
                            }));
                            
                            ret = CMD_DIR(1);

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_DIR Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 10;
                            }
                            break;
                        /** 回原點 */
                        case 10:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "回原點";
                            }));                            

                            ret = CMD_MSP(1);

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MSP Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 20;
                            }
                            break;
                        /** Xsc */
                        case 20:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "移動Xsc";
                            }));

                            ret = CMD_MRS(Convert.ToInt32(double.Parse(Xsc_txt.Text) / 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MRS Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 21;
                            }
                            break;
                        /** Xts */
                        case 21:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "移動Xts";
                            }));                            

                            ret = CMD_MLS(Convert.ToInt32(double.Parse(Xts_txt.Text) / 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MLS Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 22;
                            }
                            break;
                        /** xAS */
                        case 22:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "移動xAS";
                            }));

                            ret = CMD_MLS(Convert.ToInt32(double.Parse(xAS_txt.Text) / 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MLS Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 30;
                            }
                            break;
                        /** 開燈  等待T2時間使LED達穩態    */
                        case 30:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "延遲時間T2";
                            }));                            

                            ret = CMD_SWL(1);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_SWL Error!");
                                iTask = 999;
                            }
                            else
                            {
                                Task.Delay(Convert.ToInt32(T2 * 1000)).Wait();
                                iTask = 40;
                            }
                            break;
                        /** Auto-scaling */
                        case 40:    //初始AG DG EXP
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "Auto-scalin";
                            }));                            

                            ret = CMD_AGN(int.Parse(AG_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_AGN Error!");
                                iTask = 999;
                                break;
                            }

                            ret = CMD_GNV(int.Parse(DG_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_GNV Error!");
                                iTask = 999;
                                break;
                            }

                            ret = CMD_ELC(int.Parse(EXP_initial_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_ELC Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 41:    //AG增加
                            AG *= 2;
                            ret = CMD_AGN(AG);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_AGN Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 42:    //DG增加
                            DG *= 2;
                            ret = CMD_GNV(DG);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_GNV Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 43:    //EXP減一半
                            EXP_init /= 2;
                            ret = CMD_ELC(EXP_init);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_ELC Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 45:    //找出目標值的EXP  
                            Task.Delay(1000).Wait();
                            List<int> sp1 = CMD_CAL();
                            Task.Delay(1000).Wait();
                            maxValue = sp1.Max();


                            if (maxValue < int.Parse(I_max_txt.Text))
                            {
                                EXP1 = EXP_init;
                                I1 = maxValue;

                                EXP_init /= 2;
                                ret = CMD_ELC(EXP_init);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_ELC Error!");
                                    iTask = 999;
                                    break;
                                }

                                Task.Delay(1000).Wait();
                                List<int> sp2 = CMD_CAL();
                                Task.Delay(1000).Wait();
                                maxValue = sp2.Max();


                                EXP2 = EXP_init;
                                I2 = maxValue;

                                /*System.Diagnostics.Debug.WriteLine("AG...." + AG.ToString());
                                System.Diagnostics.Debug.WriteLine("DG...." + DG.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP...." + EXP.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP1...." + EXP1.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP2...." + EXP2.ToString());
                                System.Diagnostics.Debug.WriteLine("I1...." + I1.ToString());
                                System.Diagnostics.Debug.WriteLine("I2...." + I2.ToString());*/
                                EXP = EXP1 + ((I_thr - I1) * ((EXP1 - EXP2) / (I1 - I2)));

                                if (EXP > int.Parse(EXP_max_txt.Text))
                                {
                                    if (DG < 128)
                                    {
                                        EXP_init = int.Parse(EXP_initial_txt.Text);
                                        ret = CMD_ELC(EXP_init);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_ELC Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        iTask = 42;
                                        break;
                                    }
                                    else if (AG < 8)
                                    {
                                        DG = int.Parse(DG_txt.Text);
                                        ret = CMD_GNV(DG);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_GNV Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        EXP_init = int.Parse(EXP_initial_txt.Text);
                                        ret = CMD_ELC(EXP_init);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_ELC Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        iTask = 41;
                                        break;
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine("調不到目標值");
                                    }
                                }
                                else
                                {
                                    ret = CMD_AGN(AG);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_AGN Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    ret = CMD_GNV(DG);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_GNV Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    ret = CMD_ELC(EXP);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_ELC Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    Task.Delay(1000).Wait();
                                    List<int> sp3 = CMD_CAL();
                                    Task.Delay(1000).Wait();

                                    CAL_cycle = 0;
                                    iTask = 50;                                    
                                }
                            }
                            else
                            {
                                iTask = 43;
                            }
                            break;
                        /** 移到第一點 */
                        case 50:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "移到第一點";
                            }));

                            ret = CMD_MLS(Convert.ToInt32((double.Parse(x1_txt.Text) - double.Parse(xAS_txt.Text))/ 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MLS Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 51;
                            }
                            break;
                        /** 開始掃描 */
                        case 51:
                            /*BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "掃描中";
                            }));*/



                            if (CAL_cycle < int.Parse(total_point_txt.Text))
                            {
                                ret = CMD_MLS(int.Parse(point_distance_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_MLS Error!");
                                    iTask = 999;
                                    break;
                                }

                                Task.Delay(1000).Wait();
                                List<int> sp4 = CMD_CAL();
                                Task.Delay(1000).Wait();

                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = CAL_cycle.ToString();
                                }));

                                ALL_POINT_CAL.Add(sp4);
                                CAL_cycle++;

                                iTask = 51;
                            }
                            else
                            {
                                iTask = 60;
                            }
                            break;
                        case 60:
                            List<double> n = new List<double>();
                            List<int> sp5 = new List<int>();

                            double lambda;
                            wl = int.Parse(wl_txt.Text);

                            for (int i = 1; i <= 1280; i++)
                            {
                                lambda = double.Parse(a0_txt.Text) + double.Parse(a1_txt.Text) * Convert.ToDouble(i) + double.Parse(a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                n.Add(lambda);
                            }

                            double num = n.OrderBy(item => Math.Abs(item - wl)).ThenBy(item => item).First(); //取最接近的數
                            int index = n.FindIndex(item => item.Equals(num)); //找到該數索引值

                            n.Clear();
                            for (int i = 0; i < int.Parse(total_point_txt.Text); i++)
                            {
                                n.Add(i + 1);
                                sp5.Add(ALL_POINT_CAL[i][index]);
                            }

                            
                            this.InvokeIfRequired(() =>
                            {
                                chart1.Series[0].Points.Clear();
                                chart1.Series[0].Points.DataBindXY(n, sp5);
                            });

                            iTask = 61;
                            break;
                        case 61:

                            if(wl != int.Parse(wl_txt.Text))
                            {
                                iTask = 60;
                            }

                            break;
                        case 999:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "錯誤!";
                            }));

                            break;
                    }
                }
                //timer1.Start();
            });
        }

        /**     印Log    */
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
        /**     印Log    */

        /**     開啟port      */
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

        /**     關閉port   */
        private void btnClose_Click(object sender, EventArgs e)
        {
            serialPort.Close();

            btnOpen.Visible = true;
            btnClose.Visible = false;
        }

        /**     資料傳送    */
        public void SerialWrite(string msg)
        {
            if (!serialPort.IsOpen) return;

            RawLog(string.Format("\r\nsend -> {0}\r\n", msg));
            serialPort.Write(msg);
        }

        /**     資料接收    */
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            byte data;
            SerialPort sp = (SerialPort)sender;

            char[] array = new char[24567];
            int i = 0;

            if (command == "$CAL#")
            {
                while (sp.BytesToRead > 0)
                {                    
                    data = (byte)sp.ReadByte();
                    CAL.Add(data);
                }

                /*if (CAL.Count > 5)
                {
                    String s = new String(new char[] { (char)CAL[CAL.Count - 5], (char)CAL[CAL.Count - 4], (char)CAL[CAL.Count - 3], (char)CAL[CAL.Count - 2], (char)CAL[CAL.Count - 1] });
                    if (s == "^EOF#")
                    {
                        show_Spectrum(CAL);
                    }
                }*/
            }
            else
            {
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
        }

        private void Recv_Clear()
        {
            recv_count = 0;
        }

        private string Recv_String()
        {
            recv_buff[recv_count] = '\0';
            string recv = new string(recv_buff);

            return recv;
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

        /** 讀入參數 */
        private void LoadSetting()
        {
            IniFile ini = GetIniFile();

            // load setting
            selPort = ini.IniReadValue("SETTING", "COMPORT");
            T1 = double.Parse(ini.IniReadValue("PARAMETER", "T1"));
            T2 = double.Parse(ini.IniReadValue("PARAMETER", "T2"));
            XSC = double.Parse(ini.IniReadValue("PARAMETER", "XSC"));
            XTS = double.Parse(ini.IniReadValue("PARAMETER", "XTS"));
            XAS = double.Parse(ini.IniReadValue("PARAMETER", "XAS"));
            X1 = double.Parse(ini.IniReadValue("PARAMETER", "X1"));
            total_points = double.Parse(ini.IniReadValue("PARAMETER", "TOTALPOINTS"));
            step = double.Parse(ini.IniReadValue("PARAMETER", "STEP"));
            DG = int.Parse(ini.IniReadValue("PARAMETER", "DG"));
            AG = int.Parse(ini.IniReadValue("PARAMETER", "AG"));
            EXP_init = int.Parse(ini.IniReadValue("PARAMETER", "EXPINITIAL"));
            EXP_max = int.Parse(ini.IniReadValue("PARAMETER", "EXPMAX"));
            I_max = int.Parse(ini.IniReadValue("PARAMETER", "IMAX"));
            I_thr = int.Parse(ini.IniReadValue("PARAMETER", "ITHR"));

            T1_txt.Text = T1.ToString();
            T2_txt.Text = T2.ToString();
            Xsc_txt.Text = XSC.ToString();
            Xts_txt.Text = XTS.ToString();
            xAS_txt.Text = XAS.ToString();
            x1_txt.Text = X1.ToString();
            total_point_txt.Text = total_points.ToString();
            point_distance_txt.Text = step.ToString();            
            DG_txt.Text = DG.ToString();
            AG_txt.Text = AG.ToString();
            EXP_initial_txt.Text = EXP_init.ToString();
            EXP_max_txt.Text = EXP_max.ToString();
            I_max_txt.Text = I_max.ToString();
            I_thr_txt.Text = I_thr.ToString();


            //test_led1 = bool.Parse(ini.IniReadValue("LED_TEST", "LED1"));
        }

        /**     存設定值    */
        private void SaveSetting()
        {
            // load setting
            if (cboPort.SelectedItem != null)
                selPort = cboPort.SelectedItem.ToString();


            IniFile ini = GetIniFile();

            // load setting
            ini.IniWriteValue("SETTING", "COMPORT", selPort);


            //ini.IniWriteValue("LED_TEST", "LED1", test_led1.ToString());
        }

        
        public void show_Spectrum(List<byte> list)
        {
            list.RemoveAt(0);    //刪除^
            list.RemoveAt(0);    //刪除s
            list.RemoveAt(0);    //刪除s
            list.RemoveAt(0);    //刪除s
            list.RemoveAt(0);    //刪除s
            list.RemoveAt(0);    //刪除#

            list.RemoveAt(list.Count - 1);    //刪除^
            list.RemoveAt(list.Count - 1);    //刪除E
            list.RemoveAt(list.Count - 1);    //刪除O
            list.RemoveAt(list.Count - 1);    //刪除F
            list.RemoveAt(list.Count - 1);    //刪除#

            string temp_byte = "";
            List<int> spectrum = new List<int>();
            for (int i = 0; i < list.Count; i++)
            {
                if (i % 2 == 1)
                {
                    int st = Convert.ToInt32(string.Concat(temp_byte, System.Convert.ToString(list[i], 2).PadLeft(8, '0')), 2); //2個byte轉2進制合併後再轉10進制                 
                    spectrum.Add(st);
                }
                temp_byte = System.Convert.ToString(list[i], 2).PadLeft(8, '0');
            }

            

            List<int> axis_x = new List<int>();
            for (int i = 1; i <= 1280; i++)
            {
                axis_x.Add(i);
            }

            this.InvokeIfRequired(() =>
            {
                chart1.Series[0].Points.Clear();
                chart1.Series[0].Points.DataBindXY(axis_x, spectrum);
            });

            command = "";
            CAL.Clear();

            if (checkBox4.Checked)
            {
                command = "$CAL#";
                SerialWrite("$CAL#");
            }

        }


        private void btnCommand_Click(object sender, EventArgs e)
        {
            String txt = txtCommand.Text;
            SerialWrite(txt);
            command = txt;            
        }

        private void btnCAL_Click(object sender, EventArgs e)
        {
            command = "$CAL#";
            txtCommand.Text = "";
            txtCommand.Text = "$CAL#";
        }

        private void btnAGN_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$AGN?#";
        }

        private void btnAGN1X_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$AGN1X#";
        }

        private void btnAGN2X_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$AGN2X#";
        }

        private void btnGNV_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$GNV?#";
        }

        private void btnGNV128_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$GNV128#";
        }

        private void btnELC_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$ELC?#";
        }

        private void btnELC1000_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$ELC1000#";
        }

        private void btnROI_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$ROI?#";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$ROI0,1280,488,20#";
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                command = "$CAL#";
                SerialWrite("$CAL#");
            }
        }
        

        private void RunMotorOrigin()
        {
            if (motor_back_to_origin())
            {
                Log("RunMotorOrigin OK.");

                this.InvokeIfRequired(() =>
                {
                    xAS_txt.Enabled = true;
                });
            }
            else
                Log("RunMotorOrigin Fail.");
        }

        private bool motor_back_to_origin()
        {
            int ret;
            int motor_dir = 1;
            int round = 0;

            // try DIR0
            ret = CMD_DIR(motor_dir);
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_DIR Error!");
                return false;
            }

            TestSleep(100);

            // run to pos 1
            ret = CMD_MSP(1);
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_MSP Error!");
                return false;
            }

            return true;
        }

        private bool motor_go_to_start()
        {
            int ret;

            ret = CMD_MRS(Convert.ToInt32(double.Parse(xAS_txt.Text) / 0.0196));
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_MRS Error!");
                return false;
            }

            return true;
        }

        private int CMD_MLS(int steps)
        {
            string cmd = string.Format("$MLS{0}#", steps);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(30000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("STP")) return CMD_RET_OK;
            if (recv.Contains("ERR")) return CMD_RET_ERR;

            return CMD_RET_ERR;
        }

        private int CMD_MRS(int steps)
        {
            string cmd = string.Format("$MRS{0}#", steps);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(30000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("STP")) return CMD_RET_OK;
            if (recv.Contains("ERR")) return CMD_RET_ERR;

            return CMD_RET_ERR;
        }

        private void TestSleep(int ms)
        {
            Log(string.Format("Sleep {0} ms", ms));
            Thread.Sleep(ms);
        }

        private int CMD_ELC(int EXP)
        {
            string cmd = string.Format("$ELC{0}#", EXP);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(5000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("OK")) return CMD_RET_OK;

            return CMD_RET_ERR;
        }

        private int CMD_AGN(int AG)
        {
            string cmd = string.Format("$AGN{0}X#", AG);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(5000)) return CMD_RET_TIMEOUT;
            string recv = Recv_String();
            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("OK")) return CMD_RET_OK;

            return CMD_RET_ERR;
        }

        private int CMD_GNV(int DG)
        {
            string cmd = string.Format("$GNV{0}#", DG);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(5000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("OK")) return CMD_RET_OK;

            return CMD_RET_ERR;
        }

        private int CMD_DIR(int dir)
        {
            string cmd = string.Format("$DIR{0}#", dir);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(1000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("OK")) return CMD_RET_OK;

            return CMD_RET_ERR;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$SWL0#";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$SWL1#";
        }

        /** 取光譜 */
        private List<int> CMD_CAL()
        {
            command = "$CAL#";
            string cmd = "$CAL#";

            Recv_Clear();
            SerialWrite(cmd);

            while (true)    //等光譜接收完
            {
                if(CAL.Count > 5)
                {
                    String s = new String(new char[] { (char)CAL[CAL.Count - 5], (char)CAL[CAL.Count - 4], (char)CAL[CAL.Count - 3], (char)CAL[CAL.Count - 2], (char)CAL[CAL.Count - 1] });
                    if (s == "^EOF#")
                    {
                        break;
                    }
                }                
            }


            CAL.RemoveAt(0);    //刪除^
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除#

            CAL.RemoveAt(CAL.Count - 1);    //刪除^
            CAL.RemoveAt(CAL.Count - 1);    //刪除E
            CAL.RemoveAt(CAL.Count - 1);    //刪除O
            CAL.RemoveAt(CAL.Count - 1);    //刪除F
            CAL.RemoveAt(CAL.Count - 1);    //刪除#

            string temp_byte = "";
            List<int> spectrum = new List<int>();
            for (int i = 0; i < CAL.Count; i++)
            {
                if (i % 2 == 1)
                {
                    int st = Convert.ToInt32(string.Concat(temp_byte, System.Convert.ToString(CAL[i], 2).PadLeft(8, '0')), 2); //2個byte轉2進制合併後再轉10進制                 
                    spectrum.Add(st);
                }
                temp_byte = System.Convert.ToString(CAL[i], 2).PadLeft(8, '0');
            }

            List<int> axis_x = new List<int>();
            for (int i = 1; i <= 1280; i++)
            {
                axis_x.Add(i);
            }

            this.InvokeIfRequired(() =>
            {
                chart1.Series[0].Points.Clear();
                chart1.Series[0].Points.DataBindXY(axis_x, spectrum);
            });

            command = "";
            CAL.Clear();

            return spectrum;
        }

        /*private void process_spectrum()
        {
            CAL.RemoveAt(0);    //刪除^
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除s
            CAL.RemoveAt(0);    //刪除#

            CAL.RemoveAt(CAL.Count - 1);    //刪除^
            CAL.RemoveAt(CAL.Count - 1);    //刪除E
            CAL.RemoveAt(CAL.Count - 1);    //刪除O
            CAL.RemoveAt(CAL.Count - 1);    //刪除F
            CAL.RemoveAt(CAL.Count - 1);    //刪除#

            string temp_byte = "";
            List<int> spectrum = new List<int>();

            for (int i = 0; i < CAL.Count; i++)
            {
                if (i % 2 == 1)
                {
                    int st = Convert.ToInt32(string.Concat(temp_byte, System.Convert.ToString(CAL[i], 2).PadLeft(8, '0')), 2); //2個byte轉2進制合併後再轉10進制                 
                    spectrum.Add(st);
                }
                temp_byte = System.Convert.ToString(CAL[i], 2).PadLeft(8, '0');
            }

            

            List<double> axis_x = new List<double>();
            for (int i = 1; i <= 1280; i++)
            {
                double lambda = double.Parse(a0_txt.Text) + double.Parse(a1_txt.Text) * Convert.ToDouble(i) + double.Parse(a2_txt.Text) * Math.Pow(Convert.ToDouble(i),2.0) + double.Parse(a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                
                if(Math.Floor(lambda) == double.Parse(wl_txt.Text))
                {
                    MULT_CAL.Add(spectrum[i]);
                }
                
                axis_x.Add(lambda);
            }

            command = "";
            CAL.Clear();
        }*/

        /*private void take_spectrum_btn_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() => TakeSpectrum());
        }

        private void TakeSpectrum()
        {
            if (take_spectrum())
            {
                Log("TakeSpectrum OK.");

                /*start_btn.InvokeIfRequired(() =>
                {
                    start_btn.Enabled = false;
                    start_step_txt.Enabled = false;
                    take_spectrum_btn.Enabled = true;
                    total_point.Enabled = true;
                    white_point.Enabled = true;
                    point_distance.Enabled = true;
                });
            }
            else
                Log("TakeSpectrum Fail.");
        }

        private bool take_spectrum()
        {
            int ret;
            ret = CMD_CAL();
            
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_CAL Error!");
                return false;
            }

            return true;
        }*/

        private void Run_MRS()
        {
            if (run_mrs())
            {
                if (cycle < int.Parse(total_point_txt.Text)){
                    //Task.Factory.StartNew(() => TakeSpectrum());
                    cycle++;
                    Log("Run_MRS OK.");
                }
                else
                {
                    List<int> axis_x = new List<int>();
                    for (int i = 1; i <= int.Parse(total_point_txt.Text); i++)
                    {
                        axis_x.Add(i);
                    }

                    System.Diagnostics.Debug.WriteLine(MULT_CAL.Count);

                    this.InvokeIfRequired(() =>
                    {
                        chart1.Series[0].Points.Clear();
                        chart1.Series[0].Points.DataBindXY(axis_x, MULT_CAL);
                    });
                }

            }
            else
                Log("Run_MRS Fail.");
        }

        private bool run_mrs()
        {
            int ret;

            ret = CMD_MRS(int.Parse(point_distance_txt.Text));
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_MRS Error!");
                return false;
            }

            return true;
        }

        /*private int CMD_CAL()
        {
            command = "$CAL#";
            string cmd = string.Format("$CAL#");

            Recv_Clear();
            SerialWrite(cmd);

            return CMD_RET_OK;
        }*/

        private int CMD_SWL(int state)
        {
            string cmd = string.Format("$SWL{0}#", state); //state ? 1 : 0

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(30000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("OK")) return CMD_RET_OK;

            return CMD_RET_ERR;
        }

        private int CMD_MSP(int pos)
        {
            string cmd = string.Format("$MSP{0}#", pos);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(30000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("POS")) return CMD_RET_OK;
            if (recv.Contains("ERR")) return CMD_RET_ERR;
            if (recv.Contains("WDIR")) return CMD_RET_WDIR;

            return CMD_RET_ERR;
        }

        private bool CMD_Timeout(int ms)
        {
            long st = DateTime.Now.Ticks;

            while (true)
            {
                long diff = (DateTime.Now.Ticks - st) / TimeSpan.TicksPerMillisecond;
                if (diff > ms)
                    return true;

                if (recv_count > 0 && recv_buff[recv_count - 1] == '#')
                    return false;
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
