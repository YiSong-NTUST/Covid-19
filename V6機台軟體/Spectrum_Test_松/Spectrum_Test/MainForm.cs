using Newtonsoft.Json;
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
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Spectrum_Test
{
    public delegate void AddLogDelegate(String myString); 

    

    public partial class MainForm : Form
    {

        public class Machine
        {
            public string id { get; set; }
            public string note { get; set; }
            public string t1 { get; set; }
            public string t2 { get; set; }
            public string wait { get; set; }
            public string a0 { get; set; }
            public string a1 { get; set; }
            public string a2 { get; set; }
            public string a3 { get; set; }
            public string mac { get; set; }
            public string analog_gain { get; set; }
            public string digital_gain { get; set; }
            public string xps { get; set; }
            public string xts { get; set; }
            public string xsc { get; set; }
            public string n_steps { get; set; }
            public string s_steps { get; set; }
            public string n_mov { get; set; }
            public string roi_ho { get; set; }
            public string roi_hc { get; set; }
            public string roi_vo { get; set; }
            public string roi_lc { get; set; }
            public string exp { get; set; }
            public string exp_max { get; set; }
            public string xhome { get; set; }
            public string baseline_start { get; set; }
            public string baseline_end { get; set; }
            public string xbacklash { get; set; }
            public string direction { get; set; }
            public string i_max { get; set; }
            public string i_thr { get; set; }
            public string auto_scaling { get; set; }
            public string sc_id { get; set; }
        }

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
        double SP_T1, SP_T2, SP_XSC, SP_XTS, SP_XAS, SP_X1, SP_total_points, SP_point_distance_steps, SP_step_distance;
        int SP_DG, SP_AG, SP_EXP_max, SP_EXP_init, SP_I_max, SP_I_thr;
        int SP_maxValue, SP_EXP, SP_EXP1, SP_I1, SP_EXP2, SP_I2;
        int CAL_cycle, SP_wl, CAL_RUN_cycle;

        double LED_T1, LED_T2, LED_total_times, LED_interval_time, LED_cycle_time;
        int LED_DG, LED_AG, LED_EXP_init, LED_EXP_max, LED_I_max, LED_I_thr;
        int LED_maxValue, LED_EXP, LED_EXP1, LED_I1, LED_EXP2, LED_I2;
        int LED_wl, LED_RUN_cycle,LED_AorB;

        char[] recv_buff = new char[24567];
        int recv_count = 0;
        int cycle = 1;

        bool flag;
        int iTask;
        int ret;
        int test_motor_round;

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
            point_sp_chart.ChartAreas[0].AxisY.Maximum = 4095;
            distance_sp_chart.ChartAreas[0].AxisY.Maximum = 4095;
            A_LED_chart.ChartAreas[0].AxisY.Maximum = 4095;
            B_LED_chart.ChartAreas[0].AxisY.Maximum = 4095;

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

            flag = false;
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

            // 馬達測試參數
            test_motor_round = int.Parse(ini.IniReadValue("MOTOR_TEST", "TEST_ROUND"));
            txt_motor_test_round.Text = test_motor_round.ToString();
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


            if (tabControl1.SelectedTab == tabControl1.TabPages["Motor_test_page"])
            {
                ini.IniWriteValue("MOTOR_TEST", "TEST_ROUND", txt_motor_test_round.Text);
            }
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

        private void btnLED_Test_Save_Click(object sender, EventArgs e)
        {
            SaveSetting();
        }

        private void btnAGN1X_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$AGN1X#";
        }

        private void btnApi_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtAPI.Text))
            {
                GetAllProducts(txtAPI.Text);
            }
            
        }

        private void btnAGN2X_Click(object sender, EventArgs e)
        {
            txtCommand.Text = "";
            txtCommand.Text = "$AGN2X#";
        }

        private void SP_test_point_distance_steps_txt_TextChanged(object sender, EventArgs e)
        {
            if(!String.IsNullOrEmpty(SP_test_point_distance_steps_txt.Text) && !String.IsNullOrEmpty(SP_test_step_distance_txt.Text))
            {
                SP_point_distance_txt.Text = (double.Parse(SP_test_point_distance_steps_txt.Text) * double.Parse(SP_test_step_distance_txt.Text)).ToString();
            }            
        }

        private void SP_test_step_distance_txt_TextChanged(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(SP_test_point_distance_steps_txt.Text) && !String.IsNullOrEmpty(SP_test_step_distance_txt.Text))
            {
                SP_point_distance_txt.Text = (double.Parse(SP_test_point_distance_steps_txt.Text) * double.Parse(SP_test_step_distance_txt.Text)).ToString();
            }
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

        private void btnMotor_Test_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() => RunTestMotor());
        }

        private void RunTestMotor()
        {
            if (runTestMotor())
                Log("RunTestMotor OK.");
            else
                Log("RunTestMotor Fail.");
        }

        private bool runTestMotor()
        {
            int ret;
            int motor_dir = 1;
            int round = 0;

            ret = CMD_QPS(1);
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_QPS 1 Error!");
                return false;
            }

            if (ret == CMD_RET_POS_ON)
            {
                Log("CMD_QPS 1 Detect ON!");
                return false;
            }

            TestSleep(100);

            ret = CMD_QPS(2);
            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
            {
                Log("CMD_QPS 2 Error!");
                return false;
            }

            if (ret == CMD_RET_POS_ON)
            {
                Log("CMD_QPS 2 Detect ON!");
                return false;
            }

            TestSleep(100);

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

            if (ret == CMD_RET_WDIR)
            {
                motor_dir = 0;
            }

            motor_dir = motor_dir == 1 ? 0 : 1;

            TestSleep(100);

            while (round < test_motor_round)
            {
                // try DIR0
                ret = CMD_DIR(motor_dir);
                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                {
                    Log("CMD_DIR Error!");
                    return false;
                }

                TestSleep(100);

                // run to pos 1
                ret = CMD_MSP(2);
                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                {
                    Log("CMD_MSP Error!");
                    return false;
                }

                if (ret == CMD_RET_WDIR)
                {
                    Log("CMD_RET_WDIR!");
                    return false;
                }

                TestSleep(100);

                motor_dir = motor_dir == 1 ? 0 : 1;

                // try DIR
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

                if (ret == CMD_RET_WDIR)
                {
                    Log("CMD_RET_WDIR!");
                    return false;
                }

                TestSleep(100);

                motor_dir = motor_dir == 1 ? 0 : 1;

                round++;
                Log(string.Format("Round {0} OK.", round));
            }
            return true;
        }

        /** 開始進行光譜測試程序 */
        private async void btnSpectrum_Test_Start_Click(object sender, EventArgs e)
        {
            dark_sp_chart.Series.Clear();
            dark_sp_chart.Titles.Clear();
            point_sp_chart.Series.Clear();
            point_sp_chart.Titles.Clear();
            distance_sp_chart.Series.Clear();
            distance_sp_chart.Titles.Clear();

            btnSpectrum_Test_Start.Enabled = false;
            btnSpectrum_Test_Stop.Enabled = true;
            flag = true;
            iTask = 0;
            CAL_RUN_cycle = 1;

            await Task.Run(() =>
            {
                while (flag)
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
                                iTask = 5;
                            }
                            break;
                        /** 設ROI */
                        case 5:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "設ROI";
                            }));

                            ret = CMD_ROI(int.Parse(SP_test_roi_ho_txt.Text), int.Parse(SP_test_roi_hc_txt.Text), int.Parse(SP_test_roi_vo_txt.Text), int.Parse(SP_test_roi_lc_txt.Text));

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

                            ret = CMD_MRS(Convert.ToInt32(double.Parse(SP_test_Xsc_txt.Text) / 0.0196));

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

                            ret = CMD_MLS(Convert.ToInt32(double.Parse(SP_test_Xts_txt.Text) / 0.0196));

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

                            ret = CMD_MLS(Convert.ToInt32(double.Parse(SP_test_xAS_txt.Text) / 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MLS Error!");
                                iTask = 999;
                            }
                            else
                            {
                                iTask = 23;
                            }
                            break;
                        /** 暗光譜 */
                        case 23:
                            List<double> dark_n = new List<double>();
                            double dark_lambda;

                            Task.Delay(1000).Wait();
                            List<int> dark_sp = CMD_CAL();

                            for (int i = 1; i <= 1280; i++)
                            {
                                dark_lambda = double.Parse(SP_test_a0_txt.Text) + double.Parse(SP_test_a1_txt.Text) * Convert.ToDouble(i) + double.Parse(SP_test_a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(SP_test_a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                dark_n.Add(dark_lambda);
                            }

                            Series dark_Srs = new Series("第" + CAL_RUN_cycle.ToString() + "次暗光譜");
                            dark_Srs.ChartType = SeriesChartType.Line;
                            dark_Srs.IsValueShownAsLabel = false;
                            /*dark_Srs.MarkerStyle = MarkerStyle.Circle;
                            dark_Srs.MarkerSize = 6;*/
                            dark_Srs.Points.DataBindXY(dark_n, dark_sp);
                            dark_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                            this.InvokeIfRequired(() =>
                            {
                                dark_sp_chart.Titles.Clear();
                                dark_sp_chart.Titles.Add("暗光譜");
                                dark_sp_chart.Series.Add(dark_Srs);
                            });

                            iTask = 30;
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
                                SP_T2 = double.Parse(SP_test_T2_txt.Text);
                                Task.Delay(Convert.ToInt32(SP_T2 * 1000)).Wait();
                                iTask = 40;
                            }
                            break;
                        /** Auto-scaling */
                        case 40:    //初始AG DG EXP
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "Auto-scalin";
                            }));

                            SP_AG = int.Parse(SP_test_AG_txt.Text);
                            ret = CMD_AGN(int.Parse(SP_test_AG_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_AGN Error!");
                                iTask = 999;
                                break;
                            }

                            SP_DG = int.Parse(SP_test_DG_txt.Text);
                            ret = CMD_GNV(int.Parse(SP_test_DG_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_GNV Error!");
                                iTask = 999;
                                break;
                            }

                            SP_EXP_init = int.Parse(SP_test_EXP_initial_txt.Text);
                            ret = CMD_ELC(int.Parse(SP_test_EXP_initial_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_ELC Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 41:    //AG增加
                            SP_AG *= 2;
                            ret = CMD_AGN(SP_AG);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_AGN Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 42:    //DG增加
                            SP_DG *= 2;
                            ret = CMD_GNV(SP_DG);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_GNV Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 43:    //EXP減一半
                            SP_EXP_init /= 2;
                            ret = CMD_ELC(SP_EXP_init);
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
                            SP_maxValue = sp1.Max();

                            if (SP_maxValue < int.Parse(SP_test_I_max_txt.Text))
                            {
                                SP_EXP1 = SP_EXP_init;
                                SP_I1 = SP_maxValue;

                                SP_EXP_init /= 2;
                                ret = CMD_ELC(SP_EXP_init);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_ELC Error!");
                                    iTask = 999;
                                    break;
                                }

                                Task.Delay(1000).Wait();
                                List<int> sp2 = CMD_CAL();
                                SP_maxValue = sp2.Max();

                                SP_EXP2 = SP_EXP_init;
                                SP_I2 = SP_maxValue;

                                SP_I_thr = int.Parse(SP_test_I_thr_txt.Text);

                                /*System.Diagnostics.Debug.WriteLine("AG...." + AG.ToString());
                                System.Diagnostics.Debug.WriteLine("DG...." + DG.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP...." + EXP.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP1...." + EXP1.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP2...." + EXP2.ToString());
                                System.Diagnostics.Debug.WriteLine("I1...." + I1.ToString());
                                System.Diagnostics.Debug.WriteLine("I2...." + I2.ToString());*/
                                SP_EXP = SP_EXP1 + ((SP_I_thr - SP_I1) * ((SP_EXP1 - SP_EXP2) / (SP_I1 - SP_I2)));


                                if (SP_EXP > int.Parse(SP_test_EXP_max_txt.Text))
                                {
                                    if (SP_DG < 128)
                                    {
                                        SP_EXP_init = int.Parse(SP_test_EXP_initial_txt.Text);
                                        ret = CMD_ELC(SP_EXP_init);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_ELC Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        iTask = 42;
                                        break;
                                    }
                                    else if (SP_AG < 8)
                                    {
                                        SP_DG = int.Parse(SP_test_DG_txt.Text);
                                        ret = CMD_GNV(SP_DG);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_GNV Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        SP_EXP_init = int.Parse(SP_test_EXP_initial_txt.Text);
                                        ret = CMD_ELC(SP_EXP_init);
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
                                    ret = CMD_AGN(SP_AG);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_AGN Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    ret = CMD_GNV(SP_DG);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_GNV Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    ret = CMD_ELC(SP_EXP);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_ELC Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    Task.Delay(1000).Wait();
                                    List<int> sp3 = CMD_CAL();

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

                            ret = CMD_MLS(Convert.ToInt32((double.Parse(SP_test_x1_txt.Text) - double.Parse(SP_test_xAS_txt.Text)) / 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MLS Error!");
                                iTask = 999;
                            }
                            else
                            {
                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "掃描中";
                                    ALL_POINT_CAL.Clear();
                                }));
                                iTask = 51;


                            }
                            break;
                        /** 開始掃描 */
                        case 51:
                            if (CAL_cycle < int.Parse(SP_test_total_point_txt.Text))
                            {
                                ret = CMD_MLS(int.Parse(SP_test_point_distance_steps_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_MLS Error!");
                                    iTask = 999;
                                    break;
                                }

                                Task.Delay(1000).Wait();
                                List<int> sp4 = CMD_CAL();

                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "掃描第 " + CAL_cycle.ToString() + " 點";
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
                            List<double> n_wl = new List<double>();
                            List<double> n_dis = new List<double>();
                            List<int> sp5 = new List<int>();

                            double lambda;

                            if (String.IsNullOrEmpty(SP_test_wl_txt.Text))
                            {
                                SP_wl = 0;
                            }
                            else
                            {
                                SP_wl = int.Parse(SP_test_wl_txt.Text);
                            }

                            for (int i = 1; i <= 1280; i++)
                            {
                                lambda = double.Parse(SP_test_a0_txt.Text) + double.Parse(SP_test_a1_txt.Text) * Convert.ToDouble(i) + double.Parse(SP_test_a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(SP_test_a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                n_wl.Add(lambda);
                            }

                            double num = n_wl.OrderBy(item => Math.Abs(item - SP_wl)).ThenBy(item => item).First(); //取最接近的數
                            int index = n_wl.FindIndex(item => item.Equals(num)); //找到該數索引值

                            n_wl.Clear();
                            n_dis.Clear();
                            for (int i = 0; i < int.Parse(SP_test_total_point_txt.Text); i++)
                            {
                                n_wl.Add(i + 1);

                                // 這行要改!!!!!!! //
                                n_dis.Add(double.Parse(SP_test_x1_txt.Text) + Convert.ToDouble(i) * double.Parse(SP_test_step_distance_txt.Text) * double.Parse(SP_test_point_distance_steps_txt.Text));
                                // 這行要改!!!!!!! //

                                sp5.Add(ALL_POINT_CAL[i][index]);
                            }

                            Series point_Srs = new Series("第" + CAL_RUN_cycle.ToString() + "次循環");
                            point_Srs.ChartType = SeriesChartType.Line;
                            point_Srs.IsValueShownAsLabel = false;
                            point_Srs.Points.DataBindXY(n_wl, sp5);
                            point_Srs.MarkerStyle = MarkerStyle.Circle;
                            point_Srs.MarkerSize = 6;
                            point_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";


                            Series dis_Srs = new Series("第" + CAL_RUN_cycle.ToString() + "次循環");
                            dis_Srs.ChartType = SeriesChartType.Line;
                            dis_Srs.IsValueShownAsLabel = false;
                            dis_Srs.Points.DataBindXY(n_dis, sp5);
                            dis_Srs.MarkerStyle = MarkerStyle.Circle;
                            dis_Srs.MarkerSize = 6;
                            dis_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                            this.InvokeIfRequired(() =>
                            {
                                point_sp_chart.Titles.Clear();
                                point_sp_chart.Titles.Add("波長 " + num.ToString() + " nm 下各點光強");
                                point_sp_chart.Series.Add(point_Srs);

                                distance_sp_chart.Titles.Clear();
                                distance_sp_chart.Titles.Add("波長 " + num.ToString() + " nm 下各位置光強");
                                distance_sp_chart.Series.Add(dis_Srs);
                            });

                            iTask = 70;
                            break;
                        /** 關燈 */
                        case 70:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "延遲時間T1";
                            }));

                            ret = CMD_SWL(0);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_SWL Error!");
                                iTask = 999;
                            }
                            else
                            {
                                SP_T1 = double.Parse(SP_test_T1_txt.Text);
                                Task.Delay(Convert.ToInt32(SP_T1 * 1000)).Wait();
                                iTask = 100;
                            }
                            break;
                        case 100:

                            if (CAL_RUN_cycle < int.Parse(SP_test_CAL_RUN_cycle_txt.Text))
                            {
                                CAL_RUN_cycle++;
                                iTask = 0;
                            }
                            else
                            {
                                BeginInvoke((Action)(() =>
                                {
                                    btnSpectrum_Test_Start.Enabled = true;
                                    btnSpectrum_Test_Stop.Enabled = false;
                                    flag = false;
                                    status_lb.Text = "完成";
                                }));
                            }
                            break;
                        case 999:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "錯誤!";
                            }));
                            break;
                        default:
                            break;
                    }
                }
                //timer1.Start();
            });
        }               

        private void btnSpectrum_Test_Stop_Click(object sender, EventArgs e)
        {
            btnSpectrum_Test_Start.Enabled = true;
            btnSpectrum_Test_Stop.Enabled = false;
            flag = false;

            status_lb.Text = "";

            chart1.Titles.Clear();
            chart1.Series[0].Points.Clear();
        }

        private async void btnLED_Test_Start_Click(object sender, EventArgs e)
        {
            A_LED_chart.Series.Clear();
            A_LED_chart.Titles.Clear();

            B_LED_chart.Series.Clear();
            B_LED_chart.Titles.Clear();

            btnLED_Test_Start.Enabled = false;
            btnLED_Test_Stop.Enabled = true;
            flag = true;
            iTask = 0;
            LED_RUN_cycle = 1;
            LED_AorB = 1;

            await Task.Run(() =>
            {
                while (flag)
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
                                iTask = 5;
                            }
                            break;
                        /** 設ROI */
                        case 5:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "設ROI";
                            }));

                            ret = CMD_ROI(int.Parse(LED_test_roi_ho_txt.Text), int.Parse(LED_test_roi_hc_txt.Text), int.Parse(LED_test_roi_vo_txt.Text), int.Parse(LED_test_roi_lc_txt.Text));

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
                        /** 移到中間 */
                        case 20:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "移至中間";
                            }));

                            ret = CMD_MRS(Convert.ToInt32(17.0 / 0.0196));

                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_MRS Error!");
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

                            if(LED_AorB == 1)
                            {
                                ret = CMD_SWL(1);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_SWL Error!");
                                    iTask = 999;
                                }
                                else
                                {
                                    LED_T2 = double.Parse(LED_test_T2_txt.Text);
                                    Task.Delay(Convert.ToInt32(LED_T2 * 1000)).Wait();
                                    iTask = 40;
                                }
                            }
                            else if (LED_AorB == 2)
                            {
                                ret = CMD_SUV(1);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_SUV Error!");
                                    iTask = 999;
                                }
                                else
                                {
                                    LED_T2 = double.Parse(LED_test_T2_txt.Text);
                                    Task.Delay(Convert.ToInt32(LED_T2 * 1000)).Wait();
                                    iTask = 40;
                                }
                            }                            
                            break;
                        /** Auto-scaling */
                        case 40:    //初始AG DG EXP
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "Auto-scalin";
                            }));

                            LED_AG = int.Parse(LED_test_AG_txt.Text);
                            ret = CMD_AGN(int.Parse(LED_test_AG_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_AGN Error!");
                                iTask = 999;
                                break;
                            }

                            LED_DG = int.Parse(LED_test_DG_txt.Text);
                            ret = CMD_GNV(int.Parse(LED_test_DG_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_GNV Error!");
                                iTask = 999;
                                break;
                            }

                            LED_EXP_init = int.Parse(LED_test_EXP_initial_txt.Text);
                            ret = CMD_ELC(int.Parse(LED_test_EXP_initial_txt.Text));
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_ELC Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 41:    //AG增加
                            LED_AG *= 2;
                            ret = CMD_AGN(LED_AG);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_AGN Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 42:    //DG增加
                            LED_DG *= 2;
                            ret = CMD_GNV(LED_DG);
                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                            {
                                Log("CMD_GNV Error!");
                                iTask = 999;
                                break;
                            }

                            iTask = 45;
                            break;
                        case 43:    //EXP減一半
                            LED_EXP_init /= 2;
                            ret = CMD_ELC(LED_EXP_init);
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
                            LED_maxValue = sp1.Max();

                            if (LED_maxValue < int.Parse(LED_test_I_max_txt.Text))
                            {
                                LED_EXP1 = LED_EXP_init;
                                LED_I1 = LED_maxValue;

                                LED_EXP_init /= 2;
                                ret = CMD_ELC(LED_EXP_init);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_ELC Error!");
                                    iTask = 999;
                                    break;
                                }

                                Task.Delay(1000).Wait();
                                List<int> sp2 = CMD_CAL();
                                LED_maxValue = sp2.Max();

                                LED_EXP2 = LED_EXP_init;
                                LED_I2 = LED_maxValue;

                                LED_I_thr = int.Parse(LED_test_I_thr_txt.Text);

                                /*System.Diagnostics.Debug.WriteLine("AG...." + AG.ToString());
                                System.Diagnostics.Debug.WriteLine("DG...." + DG.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP...." + EXP.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP1...." + EXP1.ToString());
                                System.Diagnostics.Debug.WriteLine("EXP2...." + EXP2.ToString());
                                System.Diagnostics.Debug.WriteLine("I1...." + I1.ToString());
                                System.Diagnostics.Debug.WriteLine("I2...." + I2.ToString());*/
                                LED_EXP = LED_EXP1 + ((LED_I_thr - LED_I1) * ((LED_EXP1 - LED_EXP2) / (LED_I1 - LED_I2)));


                                if (LED_EXP > int.Parse(LED_test_EXP_max_txt.Text))
                                {
                                    if (LED_DG < 128)
                                    {
                                        LED_EXP_init = int.Parse(LED_test_EXP_initial_txt.Text);
                                        ret = CMD_ELC(LED_EXP_init);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_ELC Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        iTask = 42;
                                        break;
                                    }
                                    else if (LED_AG < 8)
                                    {
                                        LED_DG = int.Parse(LED_test_DG_txt.Text);
                                        ret = CMD_GNV(LED_DG);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            Log("CMD_GNV Error!");
                                            iTask = 999;
                                            break;
                                        }

                                        LED_EXP_init = int.Parse(LED_test_EXP_initial_txt.Text);
                                        ret = CMD_ELC(LED_EXP_init);
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
                                    ret = CMD_AGN(LED_AG);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_AGN Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    ret = CMD_GNV(LED_DG);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_GNV Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    ret = CMD_ELC(LED_EXP);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        Log("CMD_ELC Error!");
                                        iTask = 999;
                                        break;
                                    }

                                    Task.Delay(1000).Wait();
                                    List<int> sp3 = CMD_CAL();

                                    LED_cycle_time = 0.0;

                                    BeginInvoke((Action)(() =>
                                    {
                                        status_lb.Text = "掃描中";
                                        ALL_POINT_CAL.Clear();
                                    }));

                                    iTask = 50;
                                }
                            }
                            else
                            {
                                iTask = 43;
                            }
                            break;
                        /** 開始測試 */
                        case 50:
                            if (LED_cycle_time < (double.Parse(LED_test_total_times_txt.Text)*60.0))
                            {
                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "掃描第 " + (LED_cycle_time / 60.0).ToString() + " 分";
                                }));

                                Task.Delay(Convert.ToInt32((double.Parse(LED_test_interval_time_txt.Text) * 60.0)*1000.0)).Wait();

                                //Task.Delay(1000).Wait();
                                List<int> sp4 = CMD_CAL();                               

                                ALL_POINT_CAL.Add(sp4);
                                LED_cycle_time += (double.Parse(LED_test_interval_time_txt.Text) * 60.0);

                                iTask = 50;
                            }
                            else
                            {
                                iTask = 60;
                            }
                            break;
                        case 60:
                            List<double> n_wl = new List<double>();
                            List<int> sp5 = new List<int>();

                            double lambda;

                            if (String.IsNullOrEmpty(LED_test_wl_txt.Text))
                            {
                               LED_wl = 0;
                            }
                            else
                            {
                                LED_wl = int.Parse(LED_test_wl_txt.Text);
                            }

                            for (int i = 1; i <= 1280; i++)
                            {
                                lambda = double.Parse(LED_test_a0_txt.Text) + double.Parse(LED_test_a1_txt.Text) * Convert.ToDouble(i) + double.Parse(LED_test_a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(LED_test_a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                n_wl.Add(lambda);
                            }

                            double num = n_wl.OrderBy(item => Math.Abs(item - LED_wl)).ThenBy(item => item).First(); //取最接近的數
                            int index = n_wl.FindIndex(item => item.Equals(num)); //找到該數索引值

                            n_wl.Clear();
                            for (int i = 0; i < ALL_POINT_CAL.Count; i++)
                            {
                                n_wl.Add((i + 1) * double.Parse(LED_test_interval_time_txt.Text));
                                sp5.Add(ALL_POINT_CAL[i][index]);
                            }

                            Series time_Srs = new Series("燈 " + LED_AorB.ToString() + " 第 " + LED_RUN_cycle.ToString() + " 次 " + LED_cycle_time.ToString() + " 分");
                            time_Srs.ChartType = SeriesChartType.Line;
                            time_Srs.IsValueShownAsLabel = false;
                            time_Srs.Points.DataBindXY(n_wl, sp5);
                            time_Srs.IsValueShownAsLabel = true;
                            time_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                            this.InvokeIfRequired(() =>
                            {
                                if (LED_AorB == 1)
                                {
                                    A_LED_chart.Titles.Clear();
                                    A_LED_chart.Titles.Add("波長 " + num.ToString() + " nm 下各時間光強");
                                    A_LED_chart.Series.Add(time_Srs);
                                }
                                else if (LED_AorB == 2)
                                {
                                    B_LED_chart.Titles.Clear();
                                    B_LED_chart.Titles.Add("波長 " + num.ToString() + " nm 下各時間光強");
                                    B_LED_chart.Series.Add(time_Srs);
                                }
                                    
                            });

                            iTask = 70;
                            break;
                        /** 關燈 */
                        case 70:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "延遲時間T1";
                            }));

                            if (LED_AorB == 1)
                            {
                                ret = CMD_SWL(0);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_SWL Error!");
                                    iTask = 999;
                                }
                                else
                                {
                                    LED_T1 = double.Parse(LED_test_T1_txt.Text);
                                    Task.Delay(Convert.ToInt32(LED_T1 * 1000)).Wait();
                                    iTask = 100;
                                }
                            }
                            else if (LED_AorB == 2)
                            {
                                ret = CMD_SUV(0);
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    Log("CMD_SUV Error!");
                                    iTask = 999;
                                }
                                else
                                {
                                    LED_T1 = double.Parse(LED_test_T1_txt.Text);
                                    Task.Delay(Convert.ToInt32(LED_T1 * 1000)).Wait();
                                    iTask = 100;
                                }
                            }                            
                            break;
                        case 100:

                            if (LED_RUN_cycle < int.Parse(LED_test_RUN_cycle_txt.Text))
                            {
                                LED_RUN_cycle++;
                                iTask = 0;
                            }
                            else if (LED_AorB == 1)
                            {
                                iTask = 0;
                                LED_RUN_cycle = 1;
                                LED_AorB = 2;
                            }
                            else if (LED_AorB == 2)
                            {
                                BeginInvoke((Action)(() =>
                                {
                                    btnLED_Test_Start.Enabled = true;
                                    btnLED_Test_Stop.Enabled = false;
                                    flag = false;
                                    status_lb.Text = "完成";
                                }));
                            }
                            break;
                        case 999:
                            BeginInvoke((Action)(() =>
                            {
                                status_lb.Text = "錯誤!";
                            }));
                            break;
                        default:
                            break;
                    }
                }
                //timer1.Start();
            });
        }

        private void btnSpectrum_Test_Save_Click(object sender, EventArgs e)
        {
            SaveSetting();
        }

        private int CMD_QPS(int pos)
        {
            string cmd = string.Format("$QPS{0}#", pos);

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(1000)) return CMD_RET_TIMEOUT;

            string recv = Recv_String();

            if (recv.Contains("NACK")) return CMD_RET_NACK;
            if (recv.Contains("ON")) return CMD_RET_POS_ON;
            if (recv.Contains("OFF")) return CMD_RET_POS_OFF;

            return CMD_RET_ERR;
        }


        private void RunMotorOrigin()
        {
            if (motor_back_to_origin())
            {
                Log("RunMotorOrigin OK.");

                this.InvokeIfRequired(() =>
                {
                    SP_test_xAS_txt.Enabled = true;
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

            ret = CMD_MRS(Convert.ToInt32(double.Parse(SP_test_xAS_txt.Text) / 0.0196));
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

        //$ROI0,1280,488,20#
        private int CMD_ROI(int roi_ho, int roi_hc, int roi_vo, int roi_lc)
        {
            string cmd = string.Format("$ROI{0},{1},{2},{3}#", roi_ho, roi_hc, roi_vo, roi_lc);
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
            CAL.Clear();
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

        private void Run_MRS()
        {
            if (run_mrs())
            {
                if (cycle < int.Parse(SP_test_total_point_txt.Text)){
                    //Task.Factory.StartNew(() => TakeSpectrum());
                    cycle++;
                    Log("Run_MRS OK.");
                }
                else
                {
                    List<int> axis_x = new List<int>();
                    for (int i = 1; i <= int.Parse(SP_test_total_point_txt.Text); i++)
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

            ret = CMD_MRS(int.Parse(SP_test_point_distance_steps_txt.Text));
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

        private int CMD_SUV(int state)
        {
            string cmd = string.Format("$SUV{0}#", state); //state ? 1 : 0

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

        private async void GetAllProducts(string keyword)
        {
            HttpClient client = new HttpClient();
            HttpResponseMessage response = await client.GetAsync("http://corona-management.cloud.kahap.com/api/get_drive.php?keyword="+keyword);
            response.EnsureSuccessStatusCode();
            string responseBody = await response.Content.ReadAsStringAsync();

            var machine = JsonConvert.DeserializeObject <List<Machine>>(responseBody); //反序列化

            // 光譜測試參數
            SP_test_T1_txt.Text = machine[0].t1.ToString();
            SP_test_T2_txt.Text = machine[0].t2.ToString();
            SP_test_Xsc_txt.Text = machine[0].xsc.ToString();
            SP_test_Xts_txt.Text = machine[0].xts.ToString();
            /*SP_test_xAS_txt.Text = SP_XAS.ToString();
            SP_test_x1_txt.Text = SP_X1.ToString();
            SP_test_total_point_txt.Text = SP_total_points.ToString();
            SP_test_point_distance_txt.Text = SP_point_distance.ToString();
            SP_test_step_distance_txt.Text = SP_step_distance.ToString();*/
            SP_test_DG_txt.Text = machine[0].digital_gain.ToString();
            SP_test_AG_txt.Text = machine[0].analog_gain.ToString();
            SP_test_EXP_initial_txt.Text = machine[0].exp.ToString();
            SP_test_EXP_max_txt.Text = machine[0].exp_max.ToString();
            SP_test_I_max_txt.Text = machine[0].i_max.ToString();
            SP_test_I_thr_txt.Text = machine[0].i_thr.ToString();
            SP_test_roi_ho_txt.Text = machine[0].roi_ho.ToString();
            SP_test_roi_hc_txt.Text = machine[0].roi_hc.ToString();
            SP_test_roi_vo_txt.Text = machine[0].roi_vo.ToString();
            SP_test_roi_lc_txt.Text = machine[0].roi_lc.ToString();
            SP_test_a0_txt.Text = machine[0].a0.ToString();
            SP_test_a1_txt.Text = machine[0].a1.ToString();
            SP_test_a2_txt.Text = machine[0].a2.ToString();
            SP_test_a3_txt.Text = machine[0].a3.ToString();
            SP_test_step_distance_txt.Text = machine[0].s_steps.ToString();
            SP_test_point_distance_steps_txt.Text = machine[0].n_mov.ToString();

            // LED測試參數
            LED_test_T1_txt.Text = machine[0].t1.ToString();
            LED_test_T2_txt.Text = machine[0].t2.ToString();
            LED_test_DG_txt.Text = machine[0].digital_gain.ToString();
            LED_test_AG_txt.Text = machine[0].analog_gain.ToString();
            LED_test_EXP_initial_txt.Text = machine[0].exp.ToString();
            LED_test_EXP_max_txt.Text = machine[0].exp_max.ToString();
            LED_test_I_max_txt.Text = machine[0].i_max.ToString();
            LED_test_I_thr_txt.Text = machine[0].i_thr.ToString();
            LED_test_roi_ho_txt.Text = machine[0].roi_ho.ToString();
            LED_test_roi_hc_txt.Text = machine[0].roi_hc.ToString();
            LED_test_roi_vo_txt.Text = machine[0].roi_vo.ToString();
            LED_test_roi_lc_txt.Text = machine[0].roi_lc.ToString();
            LED_test_a0_txt.Text = machine[0].a0.ToString();
            LED_test_a1_txt.Text = machine[0].a1.ToString();
            LED_test_a2_txt.Text = machine[0].a2.ToString();
            LED_test_a3_txt.Text = machine[0].a3.ToString();
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
