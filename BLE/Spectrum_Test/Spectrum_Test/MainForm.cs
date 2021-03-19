using SmartTagTool;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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

        bool test_led1, test_led2, test_led3;
        int test_led_sec;

        int test_motor_round;

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

            test_led1 = bool.Parse(ini.IniReadValue("LED_TEST", "LED1"));
            test_led2 = bool.Parse(ini.IniReadValue("LED_TEST", "LED2"));
            test_led3 = bool.Parse(ini.IniReadValue("LED_TEST", "LED3"));

            test_led_sec = int.Parse(ini.IniReadValue("LED_TEST", "SECONDS"));
            
            test_motor_round = int.Parse(ini.IniReadValue("MOTOR_TEST", "TEST_ROUND"));
        }

        private void SaveSetting()
        {
            // load setting
            if (cboPort.SelectedItem != null)
                selPort = cboPort.SelectedItem.ToString();

            test_led1 = chk_test_led1.Checked;
            test_led2 = chk_test_led2.Checked;
            test_led3 = chk_test_led3.Checked;

            test_led_sec = int.Parse(txt_test_led_time.Text);


            IniFile ini = GetIniFile();

            // load setting
            ini.IniWriteValue("SETTING", "COMPORT", selPort);

            
            ini.IniWriteValue("LED_TEST", "LED1", test_led1.ToString());
            ini.IniWriteValue("LED_TEST", "LED2", test_led2.ToString());
            ini.IniWriteValue("LED_TEST", "LED3", test_led3.ToString());

            ini.IniWriteValue("LED_TEST", "SECONDS", test_led_sec.ToString());

            ini.IniWriteValue("MOTOR_TEST", "TEST_ROUND", test_motor_round.ToString());
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            selPort = "";
            this.logDelegate = new AddLogDelegate(_log);
            this.rawlogDelegate = new AddLogDelegate(_rawlog);

            LoadSetting();


            loadPorts();

            chk_test_led1.Checked = test_led1;
            chk_test_led2.Checked = test_led2;
            chk_test_led3.Checked = test_led3;

            txt_test_led_time.Text = test_led_sec.ToString();

            txt_motor_test_round.Text = test_motor_round.ToString();

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

        private void btnClose_Click(object sender, EventArgs e)
        {

            serialPort.Close();

            btnOpen.Visible = true;
            btnClose.Visible = false;
        }

        private void SerialWrite(string msg)
        {
            if (!serialPort.IsOpen) return;

            RawLog(string.Format("\r\nsend -> {0}\r\n", msg));
            serialPort.Write(msg);
        }

        private void btnWSL0_Click(object sender, EventArgs e)
        {
            SerialWrite("$SWL0#");
        }

        private void btnWSL1_Click(object sender, EventArgs e)
        {
            SerialWrite("$SWL1#");
        }

        private void btnSUV0_Click(object sender, EventArgs e)
        {
            SerialWrite("$SUV0#");
        }

        private void btnSUV1_Click(object sender, EventArgs e)
        {
            SerialWrite("$SUV1#");
        }

        private void btnSaveSetting_Click(object sender, EventArgs e)
        {
            SaveSetting();
        }

        private void TestSleep(int ms)
        {
            Log(string.Format("Sleep {0} ms", ms));
            Thread.Sleep(ms);
        }

        private void btnDIR0_Click(object sender, EventArgs e)
        {
            SerialWrite("$DIR0#");
        }

        private void btnDIR1_Click(object sender, EventArgs e)
        {
            SerialWrite("$DIR1#");
        }

        private void btnMSP_Click(object sender, EventArgs e)
        {
            try
            {
                if (rdoPos1.Checked)
                    SerialWrite("$MSP1#");
                else
                    SerialWrite("$MSP2#");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btnQPS1_Click(object sender, EventArgs e)
        {
            SerialWrite("$QPS1#");
        }

        private void btnQPS2_Click(object sender, EventArgs e)
        {
            SerialWrite("$QPS2#");
        }

        private void btnMRS_Click(object sender, EventArgs e)
        {
            try
            {
                int steps = int.Parse(txtMotorSteps.Text);
                SerialWrite(string.Format("$MRS{0}#",steps));
            }
            catch(Exception ex)
            {
                Log(ex.ToString());
            }
            
        }

        private bool CMD_Timeout(int ms)
        {
            long st = DateTime.Now.Ticks;

            while(true)
            {
                long diff = (DateTime.Now.Ticks - st) / TimeSpan.TicksPerMillisecond;
                if (diff > ms)
                    return true;

                if (recv_count>0 && recv_buff[recv_count - 1] == '#')
                    return false;
            }
            
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

        private void btnMLS_Click(object sender, EventArgs e)
        {
            try
            {
                int steps = int.Parse(txtMotorSteps.Text);
                SerialWrite(string.Format("$MLS{0}#", steps));
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
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

                round++;
                Log(string.Format("Round {0} OK.", round));
            }

            

            return true;
        }

        private void RunTestMotor()
        {
            if (runTestMotor())
                Log("RunTestMotor OK.");
            else
                Log("RunTestMotor Fail.");
        }

        private void btnMotor_Test_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() => RunTestMotor());
        }

        private void RunMRS1()
        {

            try
            {
                int i;

                int steps = int.Parse(txtMotorSteps.Text);
                int count = int.Parse(txtMotorCount.Text);
                int delay = int.Parse(txtMotorDelay.Text);

                for(i=0;i< count;i++)
                {
                    CMD_MRS(steps);
                    Thread.Sleep(delay);
                }

                Log("RunMRS1 OK.");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void RunMLS1()
        {

            try
            {
                int i;

                int steps = int.Parse(txtMotorSteps.Text);
                int count = int.Parse(txtMotorCount.Text);
                int delay = int.Parse(txtMotorDelay.Text);

                for (i = 0; i < count; i++)
                {
                    CMD_MLS(steps);
                    Thread.Sleep(delay);
                }

                Log("RunMLS1 OK.");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btnMRS1_Click(object sender, EventArgs e)
        {

            Task.Factory.StartNew(() => RunMRS1());
        }

        private void btnMLS1_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() => RunMLS1());
        }

        private void RunTestLED()
        {
            if (test_led1)
            {
                SerialWrite("$SWL1#");
                TestSleep(test_led_sec*1000);
                SerialWrite("$SWL0#");
                TestSleep(test_led_sec * 1000);
            }

            if (test_led2)
            {
                SerialWrite("$SUV1#");
                TestSleep(test_led_sec * 1000);
                SerialWrite("$SUV0#");
                TestSleep(test_led_sec * 1000);
            }


            Log("RunTestLED Completed.");

            //if (test_led3)
            //{
            //    SerialWrite("$SWL1#");
            //    Thread.Sleep(test_led_sec);
            //    SerialWrite("$SWL0#");
            //}
        }

        private void btnLED_Test_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() => RunTestLED());
        }
    }
}
