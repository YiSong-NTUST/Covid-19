using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
using Excel = Microsoft.Office.Interop.Excel;

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
        public string command = "";

        public List<List<List<int>>> ALL_A_POINT_CAL = new List<List<List<int>>>();
        public List<List<int>> A_POINT_CAL = new List<List<int>>();
        public List<List<List<int>>> ALL_B_POINT_CAL = new List<List<List<int>>>();
        public List<List<int>> B_POINT_CAL = new List<List<int>>();

        public List<List<List<int>>> ALL_A_LED_CAL = new List<List<List<int>>>();
        public List<List<int>> A_LED_CAL = new List<List<int>>();
        public List<List<List<int>>> ALL_B_LED_CAL = new List<List<List<int>>>();
        public List<List<int>> B_LED_CAL = new List<List<int>>();
        public List<byte> CAL = new List<byte>();
        public List<int> MULT_CAL = new List<int>();

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
        double T1, T2;
        int DG, AG, EXP_init, I_thr;

        int SP_maxValue, SP_EXP, SP_EXP1, SP_I1, SP_EXP2, SP_I2;
        int CAL_cycle, SP_wl, CAL_RUN_cycle;

        double LED_cycle_time;
        int LED_maxValue, LED_EXP, LED_EXP1, LED_I1, LED_EXP2, LED_I2;
        int LED_wl, LED_RUN_cycle;
        
        int LED_AorB;

        double dark;

        char[] recv_buff = new char[24567];
        int recv_count = 0;
        int cycle = 1;

        bool flag;
        int iTask;
        int ret;
        int test_motor_round;

        private CancellationTokenSource cts;
        private CancellationToken token;


        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            selPort = "";
                                  
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

        /** LED 測試記錄下來的光譜 */
        /*private void LED_test_wl_txt_TextChanged(object sender, EventArgs e)
        {
            double lambda;
            List<double> n_wl = new List<double>();
            List<int> LED_sp = new List<int>();

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


            if (ALL_A_LED_CAL.Count > 0)
            {
                n_wl = new List<double>();
                LED_sp = new List<int>();

                this.InvokeIfRequired(() =>
                {
                    A_LED_chart.Series.Clear();
                });

                for (int i = 0; i < ALL_A_LED_CAL.Count; i++)
                {
                    n_wl = new List<double>();
                    LED_sp = new List<int>();

                    for (int j = 0; j < ALL_A_LED_CAL[i].Count; j++)
                    {
                        n_wl.Add(j * double.Parse(LED_test_interval_time_txt.Text));
                        LED_sp.Add(ALL_A_LED_CAL[i][j][index]);
                    }
                    
                    Series time_Srs = new Series("燈 " + LED_AorB.ToString() + " 第 " + i.ToString() + " 次 ");
                    time_Srs.ChartType = SeriesChartType.Line;
                    time_Srs.IsValueShownAsLabel = false;
                    time_Srs.Points.DataBindXY(n_wl, LED_sp);
                    time_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                    this.InvokeIfRequired(() =>
                    {
                        A_LED_chart.ChartAreas[0].AxisX.Minimum = 0;
                        A_LED_chart.ChartAreas[0].AxisX.Maximum = n_wl[n_wl.Count - 1];
                        A_LED_chart.ChartAreas[0].AxisX.Interval = double.Parse(LED_test_interval_time_txt.Text);

                        A_LED_chart.Titles.Clear();
                        A_LED_chart.Titles.Add("A燈 波長 " + num.ToString() + " nm 下各時間光強");
                        A_LED_chart.Series.Add(time_Srs);                        
                    });

                }               
            }
                
            if (ALL_B_LED_CAL.Count > 0)
            {
                n_wl.Clear();
                LED_sp.Clear();

                this.InvokeIfRequired(() =>
                {
                    B_LED_chart.Series.Clear();
                });

                for (int i = 0; i < ALL_B_LED_CAL.Count; i++)
                {
                    n_wl.Clear();
                    LED_sp.Clear();
                    for (int j = 0; j < ALL_B_LED_CAL[i].Count; j++)
                    {
                        n_wl.Add(j * double.Parse(LED_test_interval_time_txt.Text));
                        LED_sp.Add(ALL_B_LED_CAL[i][j][index]);
                    }

                    Series time_Srs = new Series("燈 " + LED_AorB.ToString() + " 第 " + i + " 次 ");
                    time_Srs.ChartType = SeriesChartType.Line;
                    time_Srs.IsValueShownAsLabel = false;
                    time_Srs.Points.DataBindXY(n_wl, LED_sp);
                    time_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                    this.InvokeIfRequired(() =>
                    {
                        B_LED_chart.ChartAreas[0].AxisX.Minimum = 0;
                        B_LED_chart.ChartAreas[0].AxisX.Maximum = n_wl[n_wl.Count - 1];
                        B_LED_chart.ChartAreas[0].AxisX.Interval = double.Parse(LED_test_interval_time_txt.Text);

                        B_LED_chart.Titles.Clear();
                        B_LED_chart.Titles.Add("B燈 波長 " + num.ToString() + " nm 下各時間光強");                        
                        B_LED_chart.Series.Add(time_Srs);
                    });
                }
            }
        }*/

        /** 光譜測試 記錄下來的光譜 */
        /* 
        private void SP_test_wl_txt_TextChanged(object sender, EventArgs e)
        {
            List<double> n = new List<double>();
            List<double> n_wl = new List<double>();
            List<double> n_dis = new List<double>();
            List<int> sp5 = new List<int>();
            List<double> baseline = new List<double>();
            List<List<double>> Sp_Dark_Baseline = new List<List<double>>();

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
                n.Add(lambda);
                n_wl.Add(lambda);
            }

            double num = n_wl.OrderBy(item => Math.Abs(item - SP_wl)).ThenBy(item => item).First(); //取最接近的數
            int index = n_wl.FindIndex(item => item.Equals(num)); //找到該數索引值

            double base_line_835 = n_wl.OrderBy(item => Math.Abs(item - double.Parse(baseline_start_txt.Text))).ThenBy(item => item).First(); //取最接近的數
            int base_line_835_index = n_wl.FindIndex(item => item.Equals(base_line_835)); //找到該數索引值

            double base_line_845 = n_wl.OrderBy(item => Math.Abs(item - double.Parse(baseline_end_txt.Text))).ThenBy(item => item).First(); //取最接近的數
            int base_line_845_index = n_wl.FindIndex(item => item.Equals(base_line_845)); //找到該數索引值

            if (ALL_A_POINT_CAL.Count > 0)
            {
                n_wl = new List<double>();
                n_dis = new List<double>();
                sp5 = new List<int>();
                Sp_Dark_Baseline = new List<List<double>>();

                this.InvokeIfRequired(() =>
                {
                    A_point_sp_chart.Series.Clear();
                    A_distance_sp_chart.Series.Clear();
                    A_reflect_sp_chart.Series.Clear();
                });

                for (int i = 0; i < ALL_A_POINT_CAL.Count; i++)
                {
                    n_wl = new List<double>();
                    n_dis = new List<double>();
                    sp5 = new List<int>();
                    Sp_Dark_Baseline = new List<List<double>>();

                    for (int j = 0; j < ALL_A_POINT_CAL[i].Count; j++)
                    {
                        n_wl.Add(j + 1);
                        n_dis.Add(double.Parse(SP_test_x1_txt.Text) + Convert.ToDouble(j) * double.Parse(SP_test_step_distance_txt.Text) * double.Parse(SP_test_point_distance_steps_txt.Text));
                        sp5.Add(ALL_A_POINT_CAL[i][j][index]);


                        double baseline_average = 0.0;
                        for (int k = base_line_835_index; k <= base_line_845_index; k++)
                        {
                            baseline_average += ALL_A_POINT_CAL[i][j][k];
                        }

                        baseline_average /= (base_line_845_index - base_line_835_index + 1);
                        Sp_Dark_Baseline.Add(ALL_A_POINT_CAL[i][j].Select(x => x - baseline_average).ToList());
                    }

                    Series point_Srs = new Series("A燈 第" + i.ToString() + "次循環");
                    point_Srs.ChartType = SeriesChartType.Line;
                    point_Srs.IsValueShownAsLabel = false;
                    point_Srs.Points.DataBindXY(n_wl, sp5);
                    point_Srs.MarkerStyle = MarkerStyle.Circle;
                    point_Srs.MarkerSize = 5;
                    point_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";


                    Series dis_Srs = new Series("A燈 第" + i.ToString() + "次循環");
                    dis_Srs.ChartType = SeriesChartType.Line;
                    dis_Srs.IsValueShownAsLabel = false;
                    dis_Srs.Points.DataBindXY(n_dis, sp5);
                    dis_Srs.MarkerStyle = MarkerStyle.Circle;
                    dis_Srs.MarkerSize = 5;
                    dis_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                    this.InvokeIfRequired(() =>
                    {
                        A_point_sp_chart.ChartAreas[0].AxisX.Minimum = n_wl[0];
                        A_point_sp_chart.ChartAreas[0].AxisX.Maximum = n_wl[n_wl.Count - 1];
                        A_point_sp_chart.ChartAreas[0].AxisY.Maximum = 4095;
                        A_point_sp_chart.ChartAreas[0].AxisY.Minimum = 0;
                        A_point_sp_chart.Titles.Clear();
                        A_point_sp_chart.Titles.Add("波長 " + num.ToString() + " nm 下各點光強");
                        A_point_sp_chart.Series.Add(point_Srs);

                        A_distance_sp_chart.ChartAreas[0].AxisX.Minimum = n_dis[0];
                        A_distance_sp_chart.ChartAreas[0].AxisX.Maximum = n_dis[n_dis.Count - 1];
                        A_distance_sp_chart.ChartAreas[0].AxisY.Maximum = 4095;
                        A_distance_sp_chart.ChartAreas[0].AxisY.Minimum = 0;
                        A_distance_sp_chart.Titles.Clear();
                        A_distance_sp_chart.Titles.Add("波長 " + num.ToString() + " nm 下各位置光強");
                        A_distance_sp_chart.Series.Add(dis_Srs);
                    });

                    for (int j = 1; j <= int.Parse(line_num_txt.Text); j++)
                    {
                        TextBox SW_txt = (TextBox)panel1.Controls.Find("SW_txt_" + j.ToString(), true)[0];
                        TextBox Line_txt = (TextBox)panel1.Controls.Find("Line_txt_" + j.ToString(), true)[0];

                        List<double> sp_ref = Sp_Dark_Baseline[int.Parse(Line_txt.Text)].Zip(Sp_Dark_Baseline[int.Parse(SW_txt.Text)], (x, y) => x / y).ToList();

                        Series ref_Srs = new Series("A燈 第" + i.ToString() + "次" + j.ToString() + "反射光譜");
                        ref_Srs.ChartType = SeriesChartType.Line;
                        ref_Srs.IsValueShownAsLabel = false;
                        ref_Srs.Points.DataBindXY(n, sp_ref);
                        ref_Srs.MarkerStyle = MarkerStyle.Circle;
                        ref_Srs.MarkerSize = 5;
                        ref_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                        this.InvokeIfRequired(() =>
                        {
                            A_reflect_sp_chart.ChartAreas[0].AxisX.Minimum = n[0];
                            A_reflect_sp_chart.ChartAreas[0].AxisX.Maximum = n[n.Count - 1];
                            A_reflect_sp_chart.ChartAreas[0].AxisY.Maximum = 1.2;
                            A_reflect_sp_chart.ChartAreas[0].AxisY.Minimum = 0;
                            A_reflect_sp_chart.Titles.Clear();
                            A_reflect_sp_chart.Titles.Add("反射光譜");
                            A_reflect_sp_chart.Series.Add(ref_Srs);
                        });
                    }
                }
            }


            if (ALL_B_POINT_CAL.Count > 0)
            {
                n_wl = new List<double>();
                n_dis = new List<double>();
                sp5 = new List<int>();
                Sp_Dark_Baseline = new List<List<double>>();

                this.InvokeIfRequired(() =>
                {
                    B_point_sp_chart.Series.Clear();
                    B_distance_sp_chart.Series.Clear();
                    B_reflect_sp_chart.Series.Clear();
                });

                for (int i = 0; i < ALL_B_POINT_CAL.Count; i++)
                {
                    n_wl = new List<double>();
                    n_dis = new List<double>();
                    sp5 = new List<int>();
                    Sp_Dark_Baseline = new List<List<double>>();

                    for (int j = 0; j < ALL_B_POINT_CAL[i].Count; j++)
                    {
                        n_wl.Add(j + 1);
                        n_dis.Add(double.Parse(SP_test_x1_txt.Text) + Convert.ToDouble(j) * double.Parse(SP_test_step_distance_txt.Text) * double.Parse(SP_test_point_distance_steps_txt.Text));
                        sp5.Add(ALL_B_POINT_CAL[i][j][index]);


                        double baseline_average = 0.0;
                        for (int k = base_line_835_index; k <= base_line_845_index; k++)
                        {
                            baseline_average += ALL_B_POINT_CAL[i][j][k];
                        }

                        baseline_average /= (base_line_845_index - base_line_835_index + 1);

                        Sp_Dark_Baseline.Add(ALL_B_POINT_CAL[i][j].Select(x => x - baseline_average).ToList());
                    }

                    Series point_Srs = new Series("B燈 第" + i.ToString() + "次循環");
                    point_Srs.ChartType = SeriesChartType.Line;
                    point_Srs.IsValueShownAsLabel = false;
                    point_Srs.Points.DataBindXY(n_wl, sp5);
                    point_Srs.MarkerStyle = MarkerStyle.Circle;
                    point_Srs.MarkerSize = 5;
                    point_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";


                    Series dis_Srs = new Series("B燈 第" + i.ToString() + "次循環");
                    dis_Srs.ChartType = SeriesChartType.Line;
                    dis_Srs.IsValueShownAsLabel = false;
                    dis_Srs.Points.DataBindXY(n_dis, sp5);
                    dis_Srs.MarkerStyle = MarkerStyle.Circle;
                    dis_Srs.MarkerSize = 5;
                    dis_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                    this.InvokeIfRequired(() =>
                    {
                        B_point_sp_chart.ChartAreas[0].AxisX.Minimum = n_wl[0];
                        B_point_sp_chart.ChartAreas[0].AxisX.Maximum = n_wl[n_wl.Count - 1];
                        B_point_sp_chart.ChartAreas[0].AxisY.Maximum = 4095;
                        B_point_sp_chart.ChartAreas[0].AxisY.Minimum = 0;
                        B_point_sp_chart.Titles.Clear();
                        B_point_sp_chart.Titles.Add("波長 " + num.ToString() + " nm 下各點光強");
                        B_point_sp_chart.Series.Add(point_Srs);

                        B_distance_sp_chart.ChartAreas[0].AxisX.Minimum = n_dis[0];
                        B_distance_sp_chart.ChartAreas[0].AxisX.Maximum = n_dis[n_dis.Count - 1];
                        B_distance_sp_chart.ChartAreas[0].AxisY.Maximum = 4095;
                        B_distance_sp_chart.ChartAreas[0].AxisY.Minimum = 0;
                        B_distance_sp_chart.Titles.Clear();
                        B_distance_sp_chart.Titles.Add("波長 " + num.ToString() + " nm 下各位置光強");
                        B_distance_sp_chart.Series.Add(dis_Srs);
                    });

                    for (int j = 1; j <= int.Parse(line_num_txt.Text); j++)
                    {
                        TextBox SW_txt = (TextBox)panel1.Controls.Find("SW_txt_" + j.ToString(), true)[0];
                        TextBox Line_txt = (TextBox)panel1.Controls.Find("Line_txt_" + j.ToString(), true)[0];

                        List<double> sp_ref = Sp_Dark_Baseline[int.Parse(Line_txt.Text)].Zip(Sp_Dark_Baseline[int.Parse(SW_txt.Text)], (x, y) => x / y).ToList();

                        Series ref_Srs = new Series("B燈 第" + i.ToString() + "次" + j.ToString() + "反射光譜");
                        ref_Srs.ChartType = SeriesChartType.Line;
                        ref_Srs.IsValueShownAsLabel = false;
                        ref_Srs.Points.DataBindXY(n, sp_ref);
                        ref_Srs.MarkerStyle = MarkerStyle.Circle;
                        ref_Srs.MarkerSize = 5;
                        ref_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                        this.InvokeIfRequired(() =>
                        {
                            B_reflect_sp_chart.ChartAreas[0].AxisX.Minimum = n[0];
                            B_reflect_sp_chart.ChartAreas[0].AxisX.Maximum = n[n.Count - 1];
                            B_reflect_sp_chart.ChartAreas[0].AxisY.Maximum = 1.2;
                            B_reflect_sp_chart.ChartAreas[0].AxisY.Minimum = 0;
                            B_reflect_sp_chart.Titles.Clear();
                            B_reflect_sp_chart.Titles.Add("反射光譜");
                            B_reflect_sp_chart.Series.Add(ref_Srs);
                        });
                    }
                }
            }
        }*/

        private void btnTaskStop_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                cts.Cancel();
            }
        }


        private async void btnMotor_Test_Click(object sender, EventArgs e)
        {
            bool motor_result =  await Motor_Test(); //true為成功 false為失敗
        }

        private async Task<bool> Motor_Test()
        {
            if (cts != null)
            {
                return false;
            }

            cts = new CancellationTokenSource();
            token = cts.Token;

            try
            {
                await Task.Run(() =>
                {
                    int ret;
                    int motor_dir = 1;
                    int round = 0;

                    BeginInvoke((Action)(() =>
                    {
                        status_lb.Text = "馬達測試開始";
                        btnMotor_Test.Enabled = false;
                    }));

                    ret = CMD_QPS(1);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        throw new Exception("錯誤");
                    }

                    if (ret == CMD_RET_POS_ON)
                    {
                        throw new Exception("錯誤");
                    }

                    Task.Delay(100, token).Wait();

                    ret = CMD_QPS(2);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        throw new Exception("錯誤");
                    }

                    if (ret == CMD_RET_POS_ON)
                    {
                        throw new Exception("錯誤");
                    }

                    Task.Delay(100, token).Wait();

                    // try DIR0
                    ret = CMD_DIR(motor_dir);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        throw new Exception("錯誤");
                    }

                    Task.Delay(100, token).Wait();

                    // run to pos 1
                    ret = CMD_MSP(1);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        throw new Exception("錯誤");
                    }

                    if (ret == CMD_RET_WDIR)
                    {
                        motor_dir = 0;
                    }

                    motor_dir = motor_dir == 1 ? 0 : 1;

                    Task.Delay(100, token).Wait();

                    while (round < int.Parse(txt_motor_test_round.Text))
                    {
                        // try DIR0
                        ret = CMD_DIR(motor_dir);
                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                        {
                            throw new Exception("錯誤");
                        }

                        Task.Delay(100, token).Wait();

                        // run to pos 1
                        ret = CMD_MSP(2);
                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                        {
                            throw new Exception("錯誤");
                        }

                        if (ret == CMD_RET_WDIR)
                        {
                            throw new Exception("錯誤");
                        }

                        Task.Delay(100, token).Wait();

                        motor_dir = motor_dir == 1 ? 0 : 1;

                        // try DIR
                        ret = CMD_DIR(motor_dir);
                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                        {
                            throw new Exception("錯誤");
                        }

                        Task.Delay(100, token).Wait();

                        // run to pos 1
                        ret = CMD_MSP(1);
                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                        {
                            throw new Exception("錯誤");
                        }

                        if (ret == CMD_RET_WDIR)
                        {
                            throw new Exception("錯誤");
                        }

                        Task.Delay(100, token).Wait();

                        motor_dir = motor_dir == 1 ? 0 : 1;

                        round++;
                        BeginInvoke((Action)(() =>
                        {
                            status_lb.Text = "第" + round.ToString() + "趟測試 OK";
                        }));

                        if (token.IsCancellationRequested == true)
                        {
                            System.Diagnostics.Debug.WriteLine("使用者已經提出取消請求");
                            token.ThrowIfCancellationRequested();
                        }
                    }                    

                    BeginInvoke((Action)(() =>
                    {
                        status_lb.Text = "完成";
                        btnMotor_Test.Enabled = true;
                    }));

                }, cts.Token);
            }
            catch (Exception ex)
            {
                if (ex.Message == "錯誤")
                {
                    BeginInvoke((Action)(() =>
                    {
                        status_lb.Text = "錯誤!";
                        btnMotor_Test.Enabled = true;
                    }));

                    cts = null;
                }
                else
                {
                    BeginInvoke((Action)(() =>
                    {
                        status_lb.Text = "停止!";
                        btnMotor_Test.Enabled = true;
                    }));

                    cts = null;
                }
                return false;
            }

            return true;
        }

        /** 開始進行光譜測試程序 */
        private async void btnSpectrum_Test_Start_Click(object sender, EventArgs e)
        {
           bool SP_result =  await Spectrum_Test();
        }

        private async Task<bool> Spectrum_Test()
        {
            if (cts != null)
            {
                return false;
            }

            cts = new CancellationTokenSource();
            token = cts.Token;

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件

            ALL_A_POINT_CAL = new List<List<List<int>>>();
            ALL_B_POINT_CAL = new List<List<List<int>>>();

            btnSpectrum_Test_Start.Enabled = false;
            flag = true;
            iTask = 0;
            CAL_RUN_cycle = 1;
            LED_AorB = 1;

            try
            {
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

                                ret = CMD_ROI(int.Parse(roi_ho_txt.Text), int.Parse(roi_hc_txt.Text), int.Parse(roi_vo_txt.Text), int.Parse(roi_lc_txt.Text));

                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
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
                                    iTask = 999;
                                }
                                else
                                {
                                    iTask = 20;
                                }
                                break;
                            /** -Xsc+Xts+xAS */
                            case 20:
                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "移動至auto scaling點";
                                }));

                                double Xsc = double.Parse(SP_test_Xsc_txt.Text);
                                double Xts = double.Parse(SP_test_Xts_txt.Text);
                                double xAS = double.Parse(SP_test_xAS_txt.Text);
                                double X_move = -Xsc + Xts + xAS;

                                if (X_move < 0.0)
                                {
                                    ret = CMD_MRS(Convert.ToInt32((-X_move) / 0.0196));
                                }
                                else
                                {
                                    ret = CMD_MLS(Convert.ToInt32(X_move / 0.0196));
                                }

                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
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

                                Task.Delay(1000, token).Wait();
                                List<int> dark_sp = CMD_CAL();

                                dark = dark_sp.Average();

                                for (int i = 1; i <= 1280; i++)
                                {
                                    dark_lambda = double.Parse(a0_txt.Text) + double.Parse(a1_txt.Text) * Convert.ToDouble(i) + double.Parse(a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                    dark_n.Add(dark_lambda);
                                }

                                Series dark_Srs = new Series("燈" + LED_AorB.ToString() + "第" + CAL_RUN_cycle.ToString() + "次暗光譜");
                                dark_Srs.ChartType = SeriesChartType.Line;
                                dark_Srs.IsValueShownAsLabel = false;
                                /*dark_Srs.MarkerStyle = MarkerStyle.Circle;
                                dark_Srs.MarkerSize = 6;*/
                                dark_Srs.Points.DataBindXY(dark_n, dark_sp);
                                dark_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                                iTask = 30;
                                break;
                            /** 開燈  等待T2時間使LED達穩態    */
                            case 30:
                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "延遲時間T2";
                                }));

                                if (LED_AorB == 1)
                                {
                                    ret = CMD_SUV(1);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T2 = double.Parse(T2_txt.Text);
                                        Task.Delay(Convert.ToInt32(T2 * 1000), token).Wait();

                                        if (CAL_RUN_cycle > 1)
                                        {
                                            CAL_cycle = 0;

                                            BeginInvoke((Action)(() =>
                                            {
                                                if (LED_AorB == 1)
                                                {
                                                    A_POINT_CAL = new List<List<int>>();
                                                }
                                                else if (LED_AorB == 2)
                                                {
                                                    B_POINT_CAL = new List<List<int>>();
                                                }
                                            }));



                                            Task.Delay(1000, token).Wait();
                                            List<int> sp4 = CMD_CAL();

                                            if (LED_AorB == 1)
                                            {
                                                A_POINT_CAL.Add(sp4);
                                            }
                                            else if (LED_AorB == 2)
                                            {
                                                B_POINT_CAL.Add(sp4);
                                            }

                                            iTask = 50;
                                        }
                                        else
                                        {
                                            iTask = 40;
                                        }
                                    }
                                }
                                else if (LED_AorB == 2)
                                {
                                    ret = CMD_SWL(1);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T2 = double.Parse(T2_txt.Text);
                                        Task.Delay(Convert.ToInt32(T2 * 1000), token).Wait();


                                        if (CAL_RUN_cycle > 1)
                                        {
                                            CAL_cycle = 0;

                                            BeginInvoke((Action)(() =>
                                            {
                                                if (LED_AorB == 1)
                                                {
                                                    A_POINT_CAL = new List<List<int>>();
                                                }
                                                else if (LED_AorB == 2)
                                                {
                                                    B_POINT_CAL = new List<List<int>>();
                                                }
                                            }));



                                            Task.Delay(1000, token).Wait();
                                            List<int> sp4 = CMD_CAL();

                                            if (LED_AorB == 1)
                                            {
                                                A_POINT_CAL.Add(sp4);
                                            }
                                            else if (LED_AorB == 2)
                                            {
                                                B_POINT_CAL.Add(sp4);
                                            }

                                            iTask = 50;
                                        }
                                        else
                                        {
                                            iTask = 40;
                                        }
                                    }
                                }

                                break;
                            /** Auto-scaling */
                            case 40:    //初始AG DG EXP
                                sw.Reset();//碼表歸零
                                sw.Start();//碼表開始計時

                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "Auto-scaling";
                                }));

                                AG = int.Parse(AG_txt.Text);
                                ret = CMD_AGN(int.Parse(AG_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    iTask = 999;
                                    break;
                                }

                                DG = int.Parse(DG_txt.Text);
                                ret = CMD_GNV(int.Parse(DG_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    iTask = 999;
                                    break;
                                }

                                EXP_init = int.Parse(EXP_initial_txt.Text);
                                ret = CMD_ELC(int.Parse(EXP_initial_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
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
                                    iTask = 999;
                                    break;
                                }

                                iTask = 45;
                                break;
                            case 45:    //找出目標值的EXP  
                                Task.Delay(1000, token).Wait();
                                List<int> sp1 = CMD_CAL();
                                SP_maxValue = sp1.Max();

                                if (SP_maxValue < int.Parse(I_max_txt.Text))
                                {
                                    SP_EXP1 = EXP_init;
                                    SP_I1 = SP_maxValue;

                                    EXP_init /= 2;
                                    ret = CMD_ELC(EXP_init);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                        break;
                                    }

                                    Task.Delay(1000, token).Wait();
                                    List<int> sp2 = CMD_CAL();
                                    SP_maxValue = sp2.Max();

                                    SP_EXP2 = EXP_init;
                                    SP_I2 = SP_maxValue;

                                    I_thr = int.Parse(I_thr_txt.Text);

                                    /*System.Diagnostics.Debug.WriteLine("AG...." + SP_AG.ToString());
                                    System.Diagnostics.Debug.WriteLine("DG...." + SP_DG.ToString());
                                    System.Diagnostics.Debug.WriteLine("EXP...." + SP_EXP.ToString());
                                    System.Diagnostics.Debug.WriteLine("EXP1...." + SP_EXP1.ToString());
                                    System.Diagnostics.Debug.WriteLine("EXP2...." + SP_EXP2.ToString());
                                    System.Diagnostics.Debug.WriteLine("I1...." + SP_I1.ToString());
                                    System.Diagnostics.Debug.WriteLine("I2...." + SP_I2.ToString());*/
                                    SP_EXP = SP_EXP1 + ((I_thr - SP_I1) * ((SP_EXP1 - SP_EXP2) / (SP_I1 - SP_I2)));


                                    if (SP_EXP > int.Parse(EXP_max_txt.Text))
                                    {
                                        if (DG < 128)
                                        {
                                            EXP_init = int.Parse(EXP_initial_txt.Text);
                                            ret = CMD_ELC(EXP_init);
                                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                            {
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
                                                iTask = 999;
                                                break;
                                            }

                                            EXP_init = int.Parse(EXP_initial_txt.Text);
                                            ret = CMD_ELC(EXP_init);
                                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                            {
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
                                            iTask = 999;
                                            break;
                                        }

                                        ret = CMD_GNV(DG);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            iTask = 999;
                                            break;
                                        }

                                        ret = CMD_ELC(SP_EXP);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            iTask = 999;
                                            break;
                                        }

                                        Task.Delay(1000, token).Wait();
                                        List<int> sp3 = CMD_CAL();

                                        sw.Stop();//碼錶停止
                                        string minutes = sw.Elapsed.Minutes.ToString();
                                        string seconds = sw.Elapsed.Seconds.ToString();

                                        CAL_cycle = 0;

                                        BeginInvoke((Action)(() =>
                                        {
                                            if (LED_AorB == 1)
                                            {
                                                A_POINT_CAL = new List<List<int>>();
                                            }
                                            else if (LED_AorB == 2)
                                            {
                                                B_POINT_CAL = new List<List<int>>();
                                            }
                                        }));



                                        Task.Delay(1000, token).Wait();
                                        List<int> sp4 = CMD_CAL();

                                        if (LED_AorB == 1)
                                        {
                                            A_POINT_CAL.Add(sp4);
                                        }
                                        else if (LED_AorB == 2)
                                        {
                                            B_POINT_CAL.Add(sp4);
                                        }

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
                                    iTask = 999;
                                }
                                else
                                {
                                    BeginInvoke((Action)(() =>
                                    {
                                        status_lb.Text = "掃描中";

                                        if (LED_AorB == 1)
                                        {
                                            A_POINT_CAL = new List<List<int>>();
                                        }
                                        else if (LED_AorB == 2)
                                        {
                                            B_POINT_CAL = new List<List<int>>();
                                        }

                                    }));

                                    /*Task.Delay(1000, token).Wait();
                                    List<int> sp4 = CMD_CAL();
                                    POINT_CAL.Add(sp4);*/

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
                                        iTask = 999;
                                        break;
                                    }

                                    Task.Delay(1000, token).Wait();
                                    List<int> sp4 = CMD_CAL();

                                    BeginInvoke((Action)(() =>
                                    {
                                        status_lb.Text = "掃描第 " + CAL_cycle.ToString() + " 點";
                                    }));

                                    if (LED_AorB == 1)
                                    {
                                        A_POINT_CAL.Add(sp4);
                                    }
                                    else if (LED_AorB == 2)
                                    {
                                        B_POINT_CAL.Add(sp4);
                                    }

                                    CAL_cycle++;

                                    iTask = 51;
                                }
                                else
                                {
                                    if (LED_AorB == 1)
                                    {
                                        ALL_A_POINT_CAL.Add(A_POINT_CAL);
                                    }
                                    else if (LED_AorB == 2)
                                    {
                                        ALL_B_POINT_CAL.Add(B_POINT_CAL);
                                    }

                                    iTask = 60;
                                }
                                break;
                            case 60:
                                List<double> n = new List<double>();
                                List<double> n_wl = new List<double>();
                                List<double> n_dis = new List<double>();
                                List<int> sp5 = new List<int>();
                                List<List<double>> Sp_Dark_Baseline = new List<List<double>>();

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
                                    lambda = double.Parse(a0_txt.Text) + double.Parse(a1_txt.Text) * Convert.ToDouble(i) + double.Parse(a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                    n.Add(lambda);
                                    n_wl.Add(lambda);
                                }

                                double num = n_wl.OrderBy(item => Math.Abs(item - SP_wl)).ThenBy(item => item).First(); //取最接近的數
                                int index = n_wl.FindIndex(item => item.Equals(num)); //找到該數索引值

                                double base_line_835 = n_wl.OrderBy(item => Math.Abs(item - double.Parse(baseline_start_txt.Text))).ThenBy(item => item).First(); //取最接近的數
                                int base_line_835_index = n_wl.FindIndex(item => item.Equals(base_line_835)); //找到該數索引值

                                double base_line_845 = n_wl.OrderBy(item => Math.Abs(item - double.Parse(baseline_end_txt.Text))).ThenBy(item => item).First(); //取最接近的數
                                int base_line_845_index = n_wl.FindIndex(item => item.Equals(base_line_845)); //找到該數索引值

                                if (LED_AorB == 1)
                                {
                                    n_wl = new List<double>();
                                    n_dis = new List<double>();
                                    sp5 = new List<int>();
                                    Sp_Dark_Baseline = new List<List<double>>();

                                    for (int i = 0; i < ALL_A_POINT_CAL[CAL_RUN_cycle - 1].Count; i++)
                                    {
                                        n_wl.Add(i + 1);
                                        n_dis.Add(double.Parse(SP_test_x1_txt.Text) + Convert.ToDouble(i) * double.Parse(SP_test_step_distance_txt.Text) * double.Parse(SP_test_point_distance_steps_txt.Text));
                                        sp5.Add(ALL_A_POINT_CAL[CAL_RUN_cycle - 1][i][index]);


                                        double baseline_average = 0.0;
                                        for (int j = base_line_835_index; j <= base_line_845_index; j++)
                                        {
                                            baseline_average += ALL_A_POINT_CAL[CAL_RUN_cycle - 1][i][j];
                                        }

                                        baseline_average /= (base_line_845_index - base_line_835_index + 1);
                                        Sp_Dark_Baseline.Add(ALL_A_POINT_CAL[CAL_RUN_cycle - 1][i].Select(x => x - baseline_average).ToList());
                                    }

                                    Series point_Srs = new Series("A燈 第" + CAL_RUN_cycle.ToString() + "次循環");
                                    point_Srs.ChartType = SeriesChartType.Line;
                                    point_Srs.IsValueShownAsLabel = false;
                                    point_Srs.Points.DataBindXY(n_wl, sp5);
                                    point_Srs.MarkerStyle = MarkerStyle.Circle;
                                    point_Srs.MarkerSize = 5;
                                    point_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";


                                    Series dis_Srs = new Series("A燈 第" + CAL_RUN_cycle.ToString() + "次循環");
                                    dis_Srs.ChartType = SeriesChartType.Line;
                                    dis_Srs.IsValueShownAsLabel = false;
                                    dis_Srs.Points.DataBindXY(n_dis, sp5);
                                    dis_Srs.MarkerStyle = MarkerStyle.Circle;
                                    dis_Srs.MarkerSize = 5;
                                    dis_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                                    double average = sp5.Average();
                                    bool low_than_average = false;
                                    List<int> sp_mini_region = new List<int>();
                                    List<int> index_mini_region = new List<int>();
                                    List<int> sp_local_minimum = new List<int>();
                                    List<int> index_local_minimum = new List<int>();

                                    for (int i = 0; i < sp5.Count; i++)
                                    {
                                        if (sp5[i] < average)
                                        {
                                            low_than_average = true;
                                        }

                                        if (low_than_average)
                                        {
                                            if (sp5[i] > average)
                                            {
                                                sp_local_minimum.Add(sp_mini_region.Min());
                                                index_local_minimum.Add(index_mini_region[sp_mini_region.IndexOf(sp_mini_region.Min())]);

                                                index_mini_region = new List<int>();
                                                sp_mini_region = new List<int>();
                                                low_than_average = false;
                                            }
                                            else
                                            {
                                                sp_mini_region.Add(sp5[i]);
                                                index_mini_region.Add(i);
                                            }
                                        }
                                    }

                                    for (int i = 1; i <= index_local_minimum.Count; i++)
                                    {
                                        double SW_dis = double.Parse(SW_dis_txt.Text);
                                        int SW_dis_to_point = Convert.ToInt16(Math.Round(SW_dis / double.Parse(SP_point_distance_txt.Text), 0, MidpointRounding.AwayFromZero));

                                        List<double> sp_ref = Sp_Dark_Baseline[index_local_minimum[i - 1]].Zip(Sp_Dark_Baseline[(index_local_minimum[i - 1] + SW_dis_to_point)], (x, y) => x / y).ToList();

                                        Series ref_Srs = new Series("A燈 第" + CAL_RUN_cycle.ToString() + "次" + i.ToString() + "反射光譜");
                                        ref_Srs.ChartType = SeriesChartType.Line;
                                        ref_Srs.IsValueShownAsLabel = false;
                                        ref_Srs.Points.DataBindXY(n, sp_ref);
                                        ref_Srs.MarkerStyle = MarkerStyle.Circle;
                                        ref_Srs.MarkerSize = 5;
                                        ref_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                                    }
                                }
                                else if (LED_AorB == 2)
                                {
                                    n_wl = new List<double>();
                                    n_dis = new List<double>();
                                    sp5 = new List<int>();
                                    Sp_Dark_Baseline = new List<List<double>>();

                                    for (int i = 0; i < ALL_B_POINT_CAL[CAL_RUN_cycle - 1].Count; i++)
                                    {
                                        n_wl.Add(i + 1);
                                        n_dis.Add(double.Parse(SP_test_x1_txt.Text) + Convert.ToDouble(i) * double.Parse(SP_test_step_distance_txt.Text) * double.Parse(SP_test_point_distance_steps_txt.Text));
                                        sp5.Add(ALL_B_POINT_CAL[CAL_RUN_cycle - 1][i][index]);

                                        double baseline_average = 0.0;
                                        for (int j = base_line_835_index; j <= base_line_845_index; j++)
                                        {
                                            baseline_average += ALL_B_POINT_CAL[CAL_RUN_cycle - 1][i][j];
                                        }

                                        baseline_average /= (base_line_845_index - base_line_835_index + 1);
                                        Sp_Dark_Baseline.Add(ALL_B_POINT_CAL[CAL_RUN_cycle - 1][i].Select(x => x - baseline_average).ToList());

                                    }

                                    Series point_Srs = new Series("B燈 第" + CAL_RUN_cycle.ToString() + "次循環");
                                    point_Srs.ChartType = SeriesChartType.Line;
                                    point_Srs.IsValueShownAsLabel = false;
                                    point_Srs.Points.DataBindXY(n_wl, sp5);
                                    point_Srs.MarkerStyle = MarkerStyle.Circle;
                                    point_Srs.MarkerSize = 5;
                                    point_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";


                                    Series dis_Srs = new Series("B燈 第" + CAL_RUN_cycle.ToString() + "次循環");
                                    dis_Srs.ChartType = SeriesChartType.Line;
                                    dis_Srs.IsValueShownAsLabel = false;
                                    dis_Srs.Points.DataBindXY(n_dis, sp5);
                                    dis_Srs.MarkerStyle = MarkerStyle.Circle;
                                    dis_Srs.MarkerSize = 5;
                                    dis_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                                    double average = sp5.Average();
                                    bool low_than_average = false;
                                    List<int> sp_mini_region = new List<int>();
                                    List<int> index_mini_region = new List<int>();
                                    List<int> sp_local_minimum = new List<int>();
                                    List<int> index_local_minimum = new List<int>();

                                    for (int i = 0; i < sp5.Count; i++)
                                    {
                                        if (sp5[i] < average)
                                        {
                                            low_than_average = true;
                                        }

                                        if (low_than_average)
                                        {
                                            if (sp5[i] > average)
                                            {
                                                sp_local_minimum.Add(sp_mini_region.Min());
                                                index_local_minimum.Add(index_mini_region[sp_mini_region.IndexOf(sp_mini_region.Min())]);

                                                index_mini_region = new List<int>();
                                                sp_mini_region = new List<int>();
                                                low_than_average = false;
                                            }
                                            else
                                            {
                                                sp_mini_region.Add(sp5[i]);
                                                index_mini_region.Add(i);
                                            }
                                        }
                                    }



                                    for (int i = 1; i <= index_local_minimum.Count; i++)
                                    {
                                        double SW_dis = double.Parse(SW_dis_txt.Text);
                                        int SW_dis_to_point = Convert.ToInt16(Math.Round(SW_dis / double.Parse(SP_point_distance_txt.Text), 0, MidpointRounding.AwayFromZero));

                                        List<double> sp_ref = Sp_Dark_Baseline[index_local_minimum[i - 1]].Zip(Sp_Dark_Baseline[(index_local_minimum[i - 1] + SW_dis_to_point)], (x, y) => x / y).ToList();

                                        Series ref_Srs = new Series("B燈 第" + CAL_RUN_cycle.ToString() + "次" + i.ToString() + "反射光譜");
                                        ref_Srs.ChartType = SeriesChartType.Line;
                                        ref_Srs.IsValueShownAsLabel = false;
                                        ref_Srs.Points.DataBindXY(n, sp_ref);
                                        ref_Srs.MarkerStyle = MarkerStyle.Circle;
                                        ref_Srs.MarkerSize = 5;
                                        ref_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                                    }
                                }

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
                                    ret = CMD_SUV(0);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T1 = double.Parse(T1_txt.Text);
                                        Task.Delay(Convert.ToInt32(T1 * 1000), token).Wait();
                                        iTask = 100;
                                    }
                                }
                                else if (LED_AorB == 2)
                                {
                                    ret = CMD_SWL(0);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T1 = double.Parse(T1_txt.Text);
                                        Task.Delay(Convert.ToInt32(T1 * 1000), token).Wait();
                                        iTask = 100;
                                    }
                                }

                                break;
                            case 100:
                                if (CAL_RUN_cycle < int.Parse(SP_test_CAL_RUN_cycle_txt.Text))
                                {
                                    CAL_RUN_cycle++;
                                    iTask = 0;
                                }
                                else if (LED_AorB == 1)
                                {
                                    iTask = 0;
                                    CAL_RUN_cycle = 1;
                                    LED_AorB = 2;
                                }
                                else if (LED_AorB == 2)
                                {
                                    BeginInvoke((Action)(() =>
                                    {
                                        btnSpectrum_Test_Start.Enabled = true;
                                        flag = false;
                                        status_lb.Text = "完成";
                                        cts = null;
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

                        if (token.IsCancellationRequested == true)
                        {
                            System.Diagnostics.Debug.WriteLine("使用者已經提出取消請求");
                            token.ThrowIfCancellationRequested();
                        }
                    }
                }, token);
            }
            catch (Exception ex)
            {

                System.Diagnostics.Debug.WriteLine(ex.Message);

                if (LED_AorB == 1)
                {
                    ret = CMD_SUV(0);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        iTask = 999;
                    }
                }
                else if (LED_AorB == 2)
                {
                    ret = CMD_SWL(0);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        iTask = 999;
                    }
                }

                BeginInvoke((Action)(() =>
                {
                    status_lb.Text = "停止!";
                    btnSpectrum_Test_Start.Enabled = true;

                    flag = false;
                }));

                cts = null;
            }

            return true;
        }

        private async void btnLED_Test_Start_Click(object sender, EventArgs e)
        {
           bool LED_result = await LED_test();

            System.Diagnostics.Debug.WriteLine(LED_result.ToString());
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            string CRF = CMD_CRF();
            string[] CRF_array = CRF.Split(new char[] { '{','}' }, StringSplitOptions.RemoveEmptyEntries);
            JObject jo = (JObject)JsonConvert.DeserializeObject("{" + CRF_array[1] + "}");

            Machine_txt.Text = jo["ID"].ToString();

            string WL = jo["WL"].ToString();
            WL = WL.Remove(0, 6);
            WL = WL.Remove(WL.Length - 4, 4);

            string[] WL_array = WL.Split(new string[] { "\",\r\n  \"" }, StringSplitOptions.RemoveEmptyEntries);

            a0_txt.Text = WL_array[0];
            a1_txt.Text = WL_array[1];
            a2_txt.Text = WL_array[2];
            a3_txt.Text = WL_array[3];

            string[] ROI = jo["ROI"].ToString().Split(',');
            roi_ho_txt.Text = ROI[0];
            roi_hc_txt.Text = ROI[1];
            roi_vo_txt.Text = ROI[2];
            roi_lc_txt.Text = ROI[3];

        }

        private string CMD_CRF()
        {
            string cmd = string.Format("$CRF0,isb.json#");

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(5000)) return "ERR";

            string recv = Recv_String();

            if (recv.Contains("NACK")) return "ERR";

            return recv;
        }

        private void btnLoadVersion_Click(object sender, EventArgs e)
        {
            VER_txt.Text = CMD_VER();
        }


        private string CMD_VER()
        {
            string cmd = string.Format("$VER#");

            Recv_Clear();
            SerialWrite(cmd);
            if (CMD_Timeout(1000)) return "ERR";

            string recv = Recv_String();

            if (recv.Contains("NACK")) return "ERR";

            return recv;
        }

        private async Task<bool> LED_test()
        {
            if (cts != null)
            {
                return false;
            }

            cts = new CancellationTokenSource();
            token = cts.Token;

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件

            ALL_A_LED_CAL = new List<List<List<int>>>();
            ALL_B_LED_CAL = new List<List<List<int>>>();

            btnLED_Test_Start.Enabled = false;
            flag = true;
            iTask = 0;
            LED_RUN_cycle = 1;
            LED_AorB = 1;

            try
            {
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

                                ret = CMD_ROI(int.Parse(roi_ho_txt.Text), int.Parse(roi_hc_txt.Text), int.Parse(roi_vo_txt.Text), int.Parse(roi_lc_txt.Text));

                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
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

                                if (LED_AorB == 1)
                                {
                                    ret = CMD_SUV(1);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T2 = double.Parse(T2_txt.Text);
                                        Task.Delay(Convert.ToInt32(T2 * 1000), token).Wait();
                                        iTask = 40;
                                    }
                                }
                                else if (LED_AorB == 2)
                                {
                                    ret = CMD_SWL(1);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T2 = double.Parse(T2_txt.Text);
                                        Task.Delay(Convert.ToInt32(T2 * 1000), token).Wait();
                                        iTask = 40;
                                    }
                                }
                                break;
                            /** 開燈  等待T2時間使LED達穩態   不做Auto scaling */
                            case 31:
                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "延遲時間T2";
                                }));

                                if (LED_AorB == 1)
                                {
                                    ret = CMD_SUV(1);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T2 = double.Parse(T2_txt.Text);
                                        Task.Delay(Convert.ToInt32(T2 * 1000), token).Wait();

                                        LED_cycle_time = 0.0;

                                        BeginInvoke((Action)(() =>
                                        {
                                            status_lb.Text = "掃描中";

                                            if (LED_AorB == 1)
                                            {
                                                A_LED_CAL = new List<List<int>>();
                                            }
                                            else if (LED_AorB == 2)
                                            {
                                                B_LED_CAL = new List<List<int>>();
                                            }
                                        }));

                                        Task.Delay(1000, token).Wait();
                                        List<int> sp4 = CMD_CAL();

                                        if (LED_AorB == 1)
                                        {
                                            A_LED_CAL.Add(sp4);
                                        }
                                        else if (LED_AorB == 2)
                                        {
                                            B_LED_CAL.Add(sp4);
                                        }

                                        iTask = 50;
                                    }
                                }
                                else if (LED_AorB == 2)
                                {
                                    ret = CMD_SWL(1);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T2 = double.Parse(T2_txt.Text);
                                        Task.Delay(Convert.ToInt32(T2 * 1000), token).Wait();

                                        LED_cycle_time = 0.0;

                                        BeginInvoke((Action)(() =>
                                        {
                                            status_lb.Text = "掃描中";

                                            if (LED_AorB == 1)
                                            {
                                                A_LED_CAL = new List<List<int>>();
                                            }
                                            else if (LED_AorB == 2)
                                            {
                                                B_LED_CAL = new List<List<int>>();
                                            }
                                        }));

                                        Task.Delay(1000, token).Wait();
                                        List<int> sp4 = CMD_CAL();

                                        if (LED_AorB == 1)
                                        {
                                            A_LED_CAL.Add(sp4);
                                        }
                                        else if (LED_AorB == 2)
                                        {
                                            B_LED_CAL.Add(sp4);
                                        }

                                        iTask = 50;
                                    }
                                }
                                break;
                            /** Auto-scaling */
                            case 40:    //初始AG DG EXP

                                sw.Reset();//碼表歸零
                                sw.Start();//碼表開始計時


                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "Auto-scaling";
                                }));

                                AG = int.Parse(AG_txt.Text);
                                ret = CMD_AGN(int.Parse(AG_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    iTask = 999;
                                    break;
                                }

                                DG = int.Parse(DG_txt.Text);
                                ret = CMD_GNV(int.Parse(DG_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
                                    iTask = 999;
                                    break;
                                }

                                EXP_init = int.Parse(EXP_initial_txt.Text);
                                ret = CMD_ELC(int.Parse(EXP_initial_txt.Text));
                                if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                {
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
                                    iTask = 999;
                                    break;
                                }

                                iTask = 45;
                                break;
                            case 45:    //找出目標值的EXP  
                                Task.Delay(1000, token).Wait();
                                List<int> sp1 = CMD_CAL();
                                LED_maxValue = sp1.Max();

                                if (LED_maxValue < int.Parse(I_max_txt.Text))
                                {
                                    LED_EXP1 = EXP_init;
                                    LED_I1 = LED_maxValue;

                                    EXP_init /= 2;
                                    ret = CMD_ELC(EXP_init);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                        break;
                                    }

                                    Task.Delay(1000, token).Wait();
                                    List<int> sp2 = CMD_CAL();
                                    LED_maxValue = sp2.Max();

                                    LED_EXP2 = EXP_init;
                                    LED_I2 = LED_maxValue;

                                    I_thr = int.Parse(I_thr_txt.Text);

                                    /*System.Diagnostics.Debug.WriteLine("AG...." + LED_AG.ToString());
                                    System.Diagnostics.Debug.WriteLine("DG...." + LED_DG.ToString());
                                    System.Diagnostics.Debug.WriteLine("EXP...." + LED_EXP.ToString());
                                    System.Diagnostics.Debug.WriteLine("EXP1...." + LED_EXP1.ToString());
                                    System.Diagnostics.Debug.WriteLine("EXP2...." + LED_EXP2.ToString());
                                    System.Diagnostics.Debug.WriteLine("I1...." + LED_I1.ToString());
                                    System.Diagnostics.Debug.WriteLine("I2...." + LED_I2.ToString());*/
                                    LED_EXP = LED_EXP1 + ((I_thr - LED_I1) * ((LED_EXP1 - LED_EXP2) / (LED_I1 - LED_I2)));
                                    //System.Diagnostics.Debug.WriteLine("EXP...." + LED_EXP.ToString());

                                    if (LED_EXP > int.Parse(EXP_max_txt.Text))
                                    {
                                        if (DG < 128)
                                        {
                                            EXP_init = int.Parse(EXP_initial_txt.Text);
                                            ret = CMD_ELC(EXP_init);
                                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                            {
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
                                                iTask = 999;
                                                break;
                                            }

                                            EXP_init = int.Parse(EXP_initial_txt.Text);
                                            ret = CMD_ELC(EXP_init);
                                            if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                            {
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
                                            iTask = 999;
                                            break;
                                        }

                                        ret = CMD_GNV(DG);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            iTask = 999;
                                            break;
                                        }

                                        ret = CMD_ELC(LED_EXP);
                                        if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                        {
                                            iTask = 999;
                                            break;
                                        }

                                        Task.Delay(1000, token).Wait();
                                        List<int> sp3 = CMD_CAL();


                                        sw.Stop();//碼錶停止
                                        string minutes = sw.Elapsed.Minutes.ToString();
                                        string seconds = sw.Elapsed.Seconds.ToString();


                                        LED_cycle_time = 0.0;

                                        BeginInvoke((Action)(() =>
                                        {
                                            status_lb.Text = "掃描中";

                                            if (LED_AorB == 1)
                                            {
                                                A_LED_CAL = new List<List<int>>();
                                            }
                                            else if (LED_AorB == 2)
                                            {
                                                B_LED_CAL = new List<List<int>>();
                                            }
                                        }));

                                        Task.Delay(1000, token).Wait();
                                        List<int> sp4 = CMD_CAL();

                                        if (LED_AorB == 1)
                                        {
                                            A_LED_CAL.Add(sp4);
                                        }
                                        else if (LED_AorB == 2)
                                        {
                                            B_LED_CAL.Add(sp4);
                                        }

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
                                if (LED_cycle_time < (double.Parse(LED_test_total_times_txt.Text) * 60.0))
                                {
                                    BeginInvoke((Action)(() =>
                                    {
                                        status_lb.Text = "掃描第 " + (LED_cycle_time / 60.0).ToString() + " 分";
                                    }));

                                    Task.Delay(Convert.ToInt32((double.Parse(LED_test_interval_time_txt.Text) * 60.0) * 1000.0), token).Wait();

                                    Task.Delay(1000, token).Wait();
                                    List<int> sp5 = CMD_CAL();

                                    if (LED_AorB == 1)
                                    {
                                        A_LED_CAL.Add(sp5);
                                    }
                                    else if (LED_AorB == 2)
                                    {
                                        B_LED_CAL.Add(sp5);
                                    }

                                    LED_cycle_time += (double.Parse(LED_test_interval_time_txt.Text) * 60.0);
                                    iTask = 50;
                                }
                                else
                                {
                                    if (LED_AorB == 1)
                                    {
                                        ALL_A_LED_CAL.Add(A_LED_CAL);
                                    }
                                    else if (LED_AorB == 2)
                                    {
                                        ALL_B_LED_CAL.Add(B_LED_CAL);
                                    }


                                    iTask = 60;
                                }
                                break;
                            case 60:
                                double lambda;
                                List<double> n_wl = new List<double>();
                                List<int> LED_sp = new List<int>();

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
                                    lambda = double.Parse(a0_txt.Text) + double.Parse(a1_txt.Text) * Convert.ToDouble(i) + double.Parse(a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                    n_wl.Add(lambda);
                                }

                                double num = n_wl.OrderBy(item => Math.Abs(item - LED_wl)).ThenBy(item => item).First(); //取最接近的數
                                int index = n_wl.FindIndex(item => item.Equals(num)); //找到該數索引值


                                if (LED_AorB == 1)
                                {
                                    n_wl = new List<double>();
                                    LED_sp = new List<int>();

                                    for (int i = 0; i < ALL_A_LED_CAL[LED_RUN_cycle - 1].Count; i++)
                                    {
                                        n_wl.Add(i * double.Parse(LED_test_interval_time_txt.Text));
                                        LED_sp.Add(ALL_A_LED_CAL[LED_RUN_cycle - 1][i][index]);
                                    }


                                    Series time_Srs = new Series("燈 " + LED_AorB.ToString() + " 第 " + LED_RUN_cycle.ToString() + " 次 " + (LED_cycle_time / 60.0) + " 分");
                                    time_Srs.ChartType = SeriesChartType.Line;
                                    time_Srs.IsValueShownAsLabel = false;
                                    time_Srs.Points.DataBindXY(n_wl, LED_sp);
                                    time_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";


                                }
                                else if (LED_AorB == 2)
                                {
                                    n_wl = new List<double>();
                                    LED_sp = new List<int>();

                                    for (int i = 0; i < ALL_B_LED_CAL[LED_RUN_cycle - 1].Count; i++)
                                    {
                                        n_wl.Add(i * double.Parse(LED_test_interval_time_txt.Text));
                                        LED_sp.Add(ALL_B_LED_CAL[LED_RUN_cycle - 1][i][index]);
                                    }

                                    Series time_Srs = new Series("燈 " + LED_AorB.ToString() + " 第 " + LED_RUN_cycle.ToString() + " 次 " + (LED_cycle_time / 60.0) + " 分");
                                    time_Srs.ChartType = SeriesChartType.Line;
                                    time_Srs.IsValueShownAsLabel = false;
                                    time_Srs.Points.DataBindXY(n_wl, LED_sp);
                                    time_Srs.ToolTip = "X: #VALX{} Y: #VALY{}";

                                }
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
                                    ret = CMD_SUV(0);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T1 = double.Parse(T1_txt.Text);
                                        Task.Delay(Convert.ToInt32(T1 * 1000), token).Wait();
                                        iTask = 100;
                                    }
                                }
                                else if (LED_AorB == 2)
                                {
                                    ret = CMD_SWL(0);
                                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                                    {
                                        iTask = 999;
                                    }
                                    else
                                    {
                                        T1 = double.Parse(T1_txt.Text);
                                        Task.Delay(Convert.ToInt32(T1 * 1000), token).Wait();
                                        iTask = 100;
                                    }
                                }
                                break;
                            case 100:
                                if (LED_RUN_cycle < int.Parse(LED_test_RUN_cycle_txt.Text))
                                {
                                    LED_RUN_cycle++;
                                    iTask = 31;
                                }
                                else if (LED_AorB == 1)
                                {
                                    iTask = 30;
                                    LED_RUN_cycle = 1;
                                    LED_AorB = 2;
                                }
                                else if (LED_AorB == 2)
                                {
                                    n_wl = new List<double>();
                                    LED_sp = new List<int>();

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
                                        lambda = double.Parse(a0_txt.Text) + double.Parse(a1_txt.Text) * Convert.ToDouble(i) + double.Parse(a2_txt.Text) * Math.Pow(Convert.ToDouble(i), 2.0) + double.Parse(a3_txt.Text) * Math.Pow(Convert.ToDouble(i), 3.0);
                                        n_wl.Add(lambda);
                                    }

                                    num = n_wl.OrderBy(item => Math.Abs(item - LED_wl)).ThenBy(item => item).First(); //取最接近的數
                                    index = n_wl.FindIndex(item => item.Equals(num)); //找到該數索引值


                                    if (ALL_A_LED_CAL.Count > 0)
                                    {
                                        n_wl = new List<double>();
                                        LED_sp = new List<int>();

                                        for (int i = 0; i < ALL_A_LED_CAL.Count; i++)
                                        {
                                            n_wl = new List<double>();
                                            LED_sp = new List<int>();

                                            for (int j = 0; j < ALL_A_LED_CAL[i].Count; j++)
                                            {
                                                n_wl.Add(j * double.Parse(LED_test_interval_time_txt.Text));
                                                LED_sp.Add(ALL_A_LED_CAL[i][j][index]);
                                            }

                                            double average = LED_sp.Average();
                                            double sumOfSquaresOfDifferences = LED_sp.Select(val => (val - average) * (val - average)).Sum();
                                            double sd = Math.Sqrt(sumOfSquaresOfDifferences / LED_sp.Count);

                                            if (sd > double.Parse(LED_STD_txt.Text))
                                            {
                                                System.Diagnostics.Debug.WriteLine(sd.ToString() + "......." + LED_STD_txt.Text);
                                                throw new Exception("錯誤");
                                            }
                                        }
                                    }

                                    if (ALL_B_LED_CAL.Count > 0)
                                    {
                                        n_wl.Clear();
                                        LED_sp.Clear();


                                        for (int i = 0; i < ALL_B_LED_CAL.Count; i++)
                                        {
                                            n_wl.Clear();
                                            LED_sp.Clear();
                                            for (int j = 0; j < ALL_B_LED_CAL[i].Count; j++)
                                            {
                                                n_wl.Add(j * double.Parse(LED_test_interval_time_txt.Text));
                                                LED_sp.Add(ALL_B_LED_CAL[i][j][index]);
                                            }

                                            double average = LED_sp.Average();
                                            double sumOfSquaresOfDifferences = LED_sp.Select(val => (val - average) * (val - average)).Sum();
                                            double sd = Math.Sqrt(sumOfSquaresOfDifferences / LED_sp.Count);

                                            if (sd > double.Parse(LED_STD_txt.Text))
                                            {
                                                System.Diagnostics.Debug.WriteLine(sd.ToString() + "......." + LED_STD_txt.Text);
                                                throw new Exception("錯誤");
                                            }
                                        }
                                    }

                                    BeginInvoke((Action)(() =>
                                    {
                                        btnLED_Test_Start.Enabled = true;
                                        flag = false;
                                        status_lb.Text = "完成";
                                        cts = null;
                                    }));
                                }
                                break;
                            case 999:
                                BeginInvoke((Action)(() =>
                                {
                                    status_lb.Text = "錯誤!";
                                    flag = false;
                                }));
                                break;
                            default:
                                break;
                        }

                        if (token.IsCancellationRequested == true)
                        {
                            System.Diagnostics.Debug.WriteLine("使用者已經提出取消請求");
                            token.ThrowIfCancellationRequested();
                        }
                    }
                }, token);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);

                if (LED_AorB == 1)
                {
                    ret = CMD_SUV(0);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        iTask = 999;
                    }
                }
                else if (LED_AorB == 2)
                {
                    ret = CMD_SWL(0);
                    if (ret == CMD_RET_TIMEOUT || ret == CMD_RET_ERR || ret == CMD_RET_NACK)
                    {
                        iTask = 999;
                    }
                }

                BeginInvoke((Action)(() =>
                {
                    status_lb.Text = "停止!";
                    btnLED_Test_Start.Enabled = true;


                    flag = false;
                }));

                cts = null;
                return false;
            }

            return true;
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

            command = "";
            CAL.Clear();

            return spectrum;
        }


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
            SP_test_Xsc_txt.Text = machine[0].xsc.ToString();
            SP_test_Xts_txt.Text = machine[0].xts.ToString();
            SP_test_step_distance_txt.Text = machine[0].s_steps.ToString();
            SP_test_point_distance_steps_txt.Text = machine[0].n_mov.ToString();
            baseline_start_txt.Text = machine[0].baseline_start.ToString();
            baseline_end_txt.Text = machine[0].baseline_end.ToString();

            T1_txt.Text = machine[0].t1.ToString();
            T2_txt.Text = machine[0].t2.ToString();
            DG_txt.Text = machine[0].digital_gain.ToString();
            AG_txt.Text = machine[0].analog_gain.ToString();
            EXP_initial_txt.Text = machine[0].exp.ToString();
            EXP_max_txt.Text = machine[0].exp_max.ToString();
            I_max_txt.Text = machine[0].i_max.ToString();
            I_thr_txt.Text = machine[0].i_thr.ToString();
            roi_ho_txt.Text = machine[0].roi_ho.ToString();
            roi_hc_txt.Text = machine[0].roi_hc.ToString();
            roi_vo_txt.Text = machine[0].roi_vo.ToString();
            roi_lc_txt.Text = machine[0].roi_lc.ToString();
            a0_txt.Text = machine[0].a0.ToString();
            a1_txt.Text = machine[0].a1.ToString();
            a2_txt.Text = machine[0].a2.ToString();
            a3_txt.Text = machine[0].a3.ToString();
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
