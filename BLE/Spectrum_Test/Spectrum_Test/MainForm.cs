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
        string selBLEPort;

        bool test_led1, test_led2, test_led3;
        int test_led_sec;

        int test_motor_round;

        int test_ble_scan_limit;
        int test_ble_pass_rssi;
        int test_ble_cmd_count;
        int test_ble_cmd_delay;
        int test_ble_loop_count;

        char[] recv_buff = new char[1024];

        char[] ble_recv_buff = new char[2048];

        string ble_recv;

        //char[] temp_recv_buff = new char[1024 * 1024];

        int recv_count = 0;
        int ble_recv_count = 0;
        //int temp_recv_count = 0;

        bool exit_test = false;

        public MainForm()
        {
            InitializeComponent();
        }


        private void DelayMs(int ms)
        {
            long st = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;

            while(true)
            {
                long diff = (DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond) - st;

                if (diff > ms) break;

                Application.DoEvents();
            }
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

        private void loadBLEPorts()
        {
            string[] ports = SerialPort.GetPortNames();

            cboBLEComport.Items.Clear();
            foreach (string portName in ports)
            {
                bool found = false;
                foreach (Object obj in cboBLEComport.Items)
                {
                    if (((string)obj).Equals(portName))
                    {
                        found = true;
                        break;
                    }

                }
                if (!found)
                {
                    cboBLEComport.Items.Add(portName);
                    if (selBLEPort != "" && selBLEPort.Equals(portName))
                    {
                        cboBLEComport.SelectedIndex = cboBLEComport.Items.Count - 1;
                    }
                }
            }
        }

        private void _rawlog(string message)
        {
            txtLog.AppendText(message);
            if (txtLog.Text.Length>0)
            {
                txtLog.SelectionStart = txtLog.Text.Length - 1; // add some logic if length is 0
                txtLog.SelectionLength = 0;
            }
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
            selBLEPort = ini.IniReadValue("SETTING", "BLECOMPORT");

            test_led1 = bool.Parse(ini.IniReadValue("LED_TEST", "LED1"));
            test_led2 = bool.Parse(ini.IniReadValue("LED_TEST", "LED2"));
            test_led3 = bool.Parse(ini.IniReadValue("LED_TEST", "LED3"));

            test_led_sec = int.Parse(ini.IniReadValue("LED_TEST", "SECONDS"));
            
            test_motor_round = int.Parse(ini.IniReadValue("MOTOR_TEST", "TEST_ROUND"));

            test_ble_scan_limit = int.Parse(ini.IniReadValue("BLE_TEST", "SCAN_RSSI_LIMIT"));
            test_ble_pass_rssi = int.Parse(ini.IniReadValue("BLE_TEST", "TEST_PASS_RSSI"));
            test_ble_cmd_count = int.Parse(ini.IniReadValue("BLE_TEST", "TEST_CMD_COUNT"));
            test_ble_cmd_delay = int.Parse(ini.IniReadValue("BLE_TEST", "TEST_CMD_DELAY"));
            test_ble_loop_count = int.Parse(ini.IniReadValue("BLE_TEST", "TEST_LOOP_COUNT"));
        }

        private void SaveSetting()
        {
            // load setting
            if (cboPort.SelectedItem != null)
                selPort = cboPort.SelectedItem.ToString();

            if (cboBLEComport.SelectedItem != null)
                selBLEPort = cboBLEComport.SelectedItem.ToString();

            test_led1 = chk_test_led1.Checked;
            test_led2 = chk_test_led2.Checked;
            test_led3 = chk_test_led3.Checked;

            test_led_sec = int.Parse(txt_test_led_time.Text);
            test_motor_round = int.Parse(txt_motor_test_round.Text);

            test_ble_scan_limit = int.Parse(txt_ble_scan_limit.Text);
            test_ble_pass_rssi = int.Parse(txt_ble_pass_rssi.Text);
            test_ble_cmd_count = int.Parse(txt_ble_test_count.Text);
            test_ble_cmd_delay = int.Parse(txt_ble_test_delay.Text);
            test_ble_loop_count = int.Parse(txt_ble_loop_count.Text);

            IniFile ini = GetIniFile();

            // load setting
            ini.IniWriteValue("SETTING", "COMPORT", selPort);
            ini.IniWriteValue("SETTING", "BLECOMPORT", selBLEPort);

            ini.IniWriteValue("LED_TEST", "LED1", test_led1.ToString());
            ini.IniWriteValue("LED_TEST", "LED2", test_led2.ToString());
            ini.IniWriteValue("LED_TEST", "LED3", test_led3.ToString());

            ini.IniWriteValue("LED_TEST", "SECONDS", test_led_sec.ToString());

            ini.IniWriteValue("MOTOR_TEST", "TEST_ROUND", test_motor_round.ToString());


            ini.IniWriteValue("BLE_TEST", "SCAN_RSSI_LIMIT", test_ble_scan_limit.ToString());
            ini.IniWriteValue("BLE_TEST", "TEST_PASS_RSSI", test_ble_pass_rssi.ToString());
            ini.IniWriteValue("BLE_TEST", "TEST_CMD_COUNT", test_ble_cmd_count.ToString());
            ini.IniWriteValue("BLE_TEST", "TEST_CMD_DELAY", test_ble_cmd_delay.ToString());
            ini.IniWriteValue("BLE_TEST", "TEST_LOOP_COUNT", test_ble_loop_count.ToString());
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            selPort = "";
            selBLEPort = "";

            this.logDelegate = new AddLogDelegate(_log);
            this.rawlogDelegate = new AddLogDelegate(_rawlog);

            LoadSetting();


            loadPorts();
            loadBLEPorts();

            chk_test_led1.Checked = test_led1;
            chk_test_led2.Checked = test_led2;
            chk_test_led3.Checked = test_led3;

            txt_test_led_time.Text = test_led_sec.ToString();

            txt_motor_test_round.Text = test_motor_round.ToString();

            txt_ble_scan_limit.Text = test_ble_scan_limit.ToString();
            txt_ble_pass_rssi.Text = test_ble_pass_rssi.ToString();
            txt_ble_test_count.Text = test_ble_cmd_count.ToString();
            txt_ble_test_delay.Text = test_ble_cmd_delay.ToString();
            txt_ble_loop_count.Text = test_ble_loop_count.ToString();
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
            recv_buff = new char[1024];
            recv_count = 0;
        }

        private string Recv_String()
        {
            recv_buff[recv_count] = '\0';

            string recv = new string(recv_buff);

            recv_buff = new char[1024];
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
                if (exit_test) break;

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


                if (exit_test) break;

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

                if (exit_test) break;

                round++;
                Log(string.Format("Round {0} OK.", round));
            }

            if (exit_test) return false;

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
            exit_test = false;
            Task.Factory.StartNew(() => RunTestMotor());
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

        private void btnBLEClose_Click(object sender, EventArgs e)
        {
            serialBLEPort.Close();

            btnBLEOpen.Visible = true;
            btnBLEClose.Visible = false;
        }

        private void btnBLEOpen_Click(object sender, EventArgs e)
        {
            string selBLEPort = (string)cboBLEComport.SelectedItem;

            if (selBLEPort == null || selBLEPort.Equals("")) return;

            serialBLEPort.PortName = selBLEPort;
            //serialPort.BaudRate = 460800;
            serialBLEPort.BaudRate = 38400;
            serialBLEPort.DataBits = 8;
            serialBLEPort.Handshake = Handshake.None;
            serialBLEPort.StopBits = StopBits.One;

            try
            {
                serialBLEPort.Open();

                serialBLEPort.DataReceived += SerialBLEPort_DataReceived;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information");
                return;
            }


            btnBLEOpen.Visible = false;
            btnBLEClose.Visible = true;
        }

        private void BLESerialWrite(string msg)
        {
            if (!serialBLEPort.IsOpen) return;
            //serialBLEPort.Close();
            //Thread.Sleep(100);
            //serialBLEPort.Open();
            //Thread.Sleep(100);

            string cmd = msg;

            //serialBLEPort.
            //RawLog(string.Format("\r\nble send -> {0}\r\n", msg));

            //int index = 0;
            //int remain = cmd.Length;

            //while (remain>0)
            //{
            //    int size = remain >= 2 ? 2 : 1;
            //    serialBLEPort.Write(cmd.Substring(index, size));
            //    remain -= size;
            //    index += size;
            //    Thread.Sleep(10);
            //}
            for (int i = 0; i < cmd.Length; i ++)
            {
                Thread.Sleep(2);
                //DelayMs(1);
                //if (cmd.Length - i > 2)
                serialBLEPort.Write(cmd.Substring(i, 1));
            }
            //serialBLEPort.Write(cmd);
        }

        private void BLE_Recv_Clear()
        {
            ble_recv_count = 0;

            ble_recv_buff = new char[2048];

            ble_recv = "";
        }

        private string BLE_Recv_String()
        {
            //ble_recv_buff[ble_recv_count] = '\0';
            //char[] temp = new char[ble_recv_count];
            //Array.Copy(ble_recv_buff, temp, ble_recv_count);

            //string recv = new string(temp);

            string temp = ble_recv;
            ble_recv = "";
            return temp;
        }

        private void SerialBLEPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            byte data;
            SerialPort sp = (SerialPort)sender;

            //char[] array = new char[4096];
            int i = 0;
            while (sp.BytesToRead > 0)
            {
                data = (byte)sp.ReadByte();

                if (data == 27) continue;
                //array[i++] = (char)data;
                //if (data == '\r') array[i++] = '\n';

                //if (ble_recv_count > 0 || data == '$')
                ble_recv_buff[ble_recv_count++] = (char)data;

                if (data == 13)
                {
                    try
                    {
                        ble_recv_buff[ble_recv_count] = (char)0x00;

                        string recv_string = new string(ble_recv_buff);
                        recv_string = recv_string.Trim((char)0x00);
                        recv_string = recv_string.Replace("uart_cli:~$ ", "");
                        recv_string = recv_string.Replace("[1;32m", "");
                        recv_string = recv_string.Replace("[1;31m", "");
                        recv_string = recv_string.Replace("[1;37m", "");
                        recv_string = recv_string.Replace("[12D", "");
                        recv_string = recv_string.Replace("[J", "");

                        RawLog(recv_string);

                        ble_recv_buff = new char[2048];
                        ble_recv_count = 0;

                        ble_recv += recv_string;
                    }
                    catch(Exception ex)
                    {

                    }
                    
                }
            }

            //array[i] = '\0';


            
            //this.Invoke(this.logDelegate, new string(array));
        }

        private void btn_ble_scan_off_Click(object sender, EventArgs e)
        {
            try
            {
                BLESerialWrite("scan off\r\n");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_devices_Click(object sender, EventArgs e)
        {
            try
            {
                BLESerialWrite("devices\r\n");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_connect_Click(object sender, EventArgs e)
        {
            if (txt_ble_uuid.Text.Equals("")) return;

            try
            {
                string cmd = string.Format("connect {0}\r\n", txt_ble_uuid.Text);
                BLESerialWrite(cmd);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_disconnect_Click(object sender, EventArgs e)
        {
            if (txt_ble_uuid.Text.Equals("")) return;

            try
            {
                string cmd = string.Format("disconnect {0}\r\n", txt_ble_uuid.Text);
                BLESerialWrite(cmd);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_services_Click(object sender, EventArgs e)
        {
            if (txt_ble_uuid.Text.Equals("")) return;

            try
            {
                string cmd = string.Format("gatt services {0}\r\n", txt_ble_uuid.Text);
                BLESerialWrite(cmd);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_tab_Click(object sender, EventArgs e)
        {
            try
            {
                BLESerialWrite("\t");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_send_Click(object sender, EventArgs e)
        {
            try
            {
                string cmd = string.Format("{0}\r\n", txt_ble_cmd.Text);
                BLESerialWrite(cmd);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_charact_Click(object sender, EventArgs e)
        {
            if (txt_ble_uuid.Text.Equals("")) return;
            if (txt_ble_service.Text.Equals("")) return;

            try
            {
                string cmd = string.Format("gatt characteristics {0} {1}\r\n", txt_ble_uuid.Text, txt_ble_service.Text);
                Thread.Sleep(100);
                BLESerialWrite(cmd);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_noti_Click(object sender, EventArgs e)
        {
            if (txt_ble_uuid.Text.Equals("")) return;
            if (txt_ble_charact.Text.Equals("")) return;

            try
            {
                string cmd = string.Format("gatt notification on {0} {1}\r\n", txt_ble_uuid.Text, 2);
                Thread.Sleep(100);
                BLESerialWrite(cmd);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }


        private byte HexToByte(string hex)
        {
            if (hex.Length > 2 || hex.Length <= 0)
            {
                hex = "0" + hex;
            }

            byte newByte = byte.Parse(hex, System.Globalization.NumberStyles.HexNumber);
            return newByte;
        }

        private void btn_ble_swl_Click(object sender, EventArgs e)
        {
            if (txt_ble_uuid.Text.Equals("")) return;
            if (txt_ble_charact.Text.Equals("")) return;

            try
            {
                cmd = string.Format("gatt write request {0} {1} {2}\r\n", txt_ble_uuid.Text, 3,"$SWL0#");
                Thread.Sleep(100);
                BLESerialWrite(cmd);
                //Thread.Sleep(1000);
                //cmd = string.Format("gatt read {0} {1}\r\n", txt_ble_uuid.Text, 2);
                ////BLE_Recv_Clear();
                //BLESerialWrite(cmd);
                

            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void btn_ble_write_Click(object sender, EventArgs e)
        {

        }

        //private bool CMD_Timeout(int ms)
        //{
        //    long st = DateTime.Now.Ticks;

        //    while (true)
        //    {
        //        long diff = (DateTime.Now.Ticks - st) / TimeSpan.TicksPerMillisecond;
        //        if (diff > ms)
        //            return true;

        //        if (recv_count > 0 && recv_buff[recv_count - 1] == '#')
        //            return false;
        //    }

        //}

        private Dictionary<string, BLE_Device> CMD_BLE_DEVICE(int rssi_limit)
        {
            Dictionary<string,BLE_Device> lst = new Dictionary<string, BLE_Device>();

            BLE_Recv_Clear();
            BLESerialWrite("devices\r\n");

            Thread.Sleep(500);

            string recv = BLE_Recv_String();
            // Device 5A:1D:DF:CA:FB:44  -80

            if (recv.Length <= 0) return lst;
            string[] ary_line = recv.Split('\r');

            foreach(string line in ary_line)
            {
                BLE_Device dev = new BLE_Device(line);

                if (dev.UUID == null || dev.UUID.Equals("")) continue;
                if (dev.RSSI == 0 || dev.RSSI < rssi_limit) continue;
                if (lst.Count > 0 && lst.ContainsKey(dev.UUID))
                    lst[dev.UUID].RSSI = dev.RSSI;
                else
                    lst.Add(dev.UUID, dev);
            }

            return lst;
            
        }

        string cmd;

        private int CMD_BLE_Connect(BLE_Device dev)
        {
            int retry = 5;
            string recv="";

            for (int i=0;i< retry;i++)
            {
                cmd = string.Format("connect {0}\r\n", dev.UUID);
                BLE_Recv_Clear();
                DelayMs(100);
                BLESerialWrite(cmd);
                DelayMs(500);

                recv = BLE_Recv_String();

                if (recv.Contains("Connected")) break;
            }

            if (!recv.Contains("Connected")) return CMD_RET_ERR;

            for (int i = 0; i < retry; i++)
            {
                // service discovery
                cmd = string.Format("gatt services {0}\r\n", dev.UUID);
                BLE_Recv_Clear();
                DelayMs(100);
                BLESerialWrite(cmd);
                DelayMs(500);

                recv = BLE_Recv_String();

                if (recv.Contains("UUID: 1 type")) break;
            }

            if (!recv.Contains("UUID: 1 type")) return CMD_RET_ERR;

            for (int i = 0; i < retry; i++)
            {
                // characteristics discovery
                cmd = string.Format("gatt characteristics {0} {1}\r\n", dev.UUID, 1);
                BLE_Recv_Clear();
                DelayMs(100);
                BLESerialWrite(cmd);
                DelayMs(500);

                recv = BLE_Recv_String();
                if (recv.Contains("Characteristic UUID: 3")) break;
            }

            if (!recv.Contains("Characteristic UUID: 3")) return CMD_RET_ERR;

            for (int i = 0; i < retry; i++)
            {
                // characteristics discovery
                cmd = string.Format("gatt notification on {0} {1}\r\n", dev.UUID, 2);
                BLE_Recv_Clear();
                DelayMs(100);
                BLESerialWrite(cmd);
                DelayMs(500);

                recv = BLE_Recv_String();
                // Type of write operation: 0x1
                if (recv.Contains("Type of write operation: 0x1")) break;
            }

            if (!recv.Contains("Type of write operation: 0x1")) return CMD_RET_ERR;

            return CMD_RET_OK;
        }

        private int CMD_BLE_TEST_CMD(BLE_Device dev)
        {
            int retry = 10;
            string recv = "";

            for (int i = 0; i < retry; i++)
            {
                cmd = string.Format("gatt write request {0} {1} {2}\r\n", dev.UUID, 3, "$SWL0#");
                BLE_Recv_Clear();
                //DelayMs(50);
                BLESerialWrite(cmd);
                BLE_Recv_Clear();
                //Thread.Sleep(2000);

                DelayMs(500);
                
                //cmd = string.Format("gatt read {0} {1}\r\n", dev.UUID, 2);
                //DelayMs(100);
                //BLESerialWrite(cmd);
                //DelayMs(2000);

                recv = BLE_Recv_String();

                if (recv.Contains("0x4F 0x4B")) break;
            }

            if (!recv.Contains("0x4F 0x4B")) return CMD_RET_ERR;

            // Notification data: 0x24 0x4F 0x4B 0x23
            int st = recv.IndexOf("Notification");
            recv = recv.Substring(st, recv.Length - st);
            string[] ary = recv.Split(' ');

            StringBuilder sb = new StringBuilder();

            sb.Append("Recv: ");

            foreach (string data in ary)
            {
                if (data.Contains("0x"))
                {
                    try
                    {
                        int index = data.IndexOf('x') + 1;
                        char raw = (char)HexToByte(data.Substring(index, data.Length - index));

                        sb.Append(raw.ToString());
                    }catch(Exception ex)
                    {

                    }
                    
                }
            }
            sb.Append("\r\n");

            Log(sb.ToString());

            DelayMs(50);

            

            return CMD_RET_OK;
        }


        private int CMD_BLE_DISCONNECT_CMD(BLE_Device dev)
        {
            int retry = 10;
            string recv = "";

            for (int i = 0; i < retry; i++)
            {
                cmd = string.Format("disconnect {0}\r\n", dev.UUID);
                BLE_Recv_Clear();
                //DelayMs(50);
                BLESerialWrite(cmd);
                BLE_Recv_Clear();
                //Thread.Sleep(2000);

                DelayMs(500);

                recv = BLE_Recv_String();

                if (recv.Contains("Disconnected")) break;
            }

            if (!recv.Contains("Disconnected")) return CMD_RET_ERR;

            return CMD_RET_OK;
        }

        private bool runTestBLE()
        {
            Dictionary<string, BLE_Device> lst_dev;
            BLE_Device target_dev;
            int loop_index = 0;

            

            //return true;

            for (loop_index=0; loop_index < test_ble_loop_count; loop_index++)
            {
                BLE_Recv_Clear();
                BLESerialWrite("scan on\r\n");
                DelayMs(1000);

                lst_dev = new Dictionary<string, BLE_Device>();

                for (int i = 0; i < 3; i++)
                {
                    Dictionary<string, BLE_Device> temp = CMD_BLE_DEVICE(test_ble_scan_limit);
                    foreach (KeyValuePair<string, BLE_Device> kvp in temp)
                    {
                        Log(string.Format("{0} {1}", kvp.Key, kvp.Value.RSSI));

                        if (kvp.Value.RSSI >= test_ble_pass_rssi)
                        {
                            if (lst_dev.ContainsKey(kvp.Key))
                                lst_dev[kvp.Key].RSSI = kvp.Value.RSSI;
                            else
                                lst_dev.Add(kvp.Key, kvp.Value);
                        }

                    }

                    if (exit_test) break;
                }

                if (lst_dev.Count <= 0)
                {
                    Log("No device found in Range.");
                    return false;
                }

                if (exit_test) break;

                target_dev = lst_dev.First().Value;

                Log(string.Format("target_dev {0} - {1}", target_dev.UUID, target_dev.RSSI.ToString()));

                //return true;
                //DelayMs(500);
                if (CMD_BLE_Connect(target_dev) == CMD_RET_OK)
                {
                    DelayMs(500);
                    if (exit_test)
                    {
                        CMD_BLE_DISCONNECT_CMD(target_dev);
                        break;
                    }

                    for (int i = 0; i < test_ble_cmd_count; i++)
                    {
                        if (CMD_BLE_TEST_CMD(target_dev) != CMD_RET_OK)
                        {
                            Log("BLE test cmd fail.");

                            BLESerialWrite(string.Format("disconnect {0}\r\n", target_dev.UUID));
                            return false;
                        }

                        Log(string.Format("BLE test {0} OK.", i + 1));

                        DelayMs(test_ble_cmd_delay);

                        if (exit_test)
                        {
                            CMD_BLE_DISCONNECT_CMD(target_dev);
                            break;
                        }
                    }

                    if (exit_test) break;

                    CMD_BLE_DISCONNECT_CMD(target_dev);

                    BLESerialWrite("scan off\r\n");

                    DelayMs(500);

                    if (exit_test) break;
                }
                else
                {
                    Log("Connect device error.");

                    CMD_BLE_DISCONNECT_CMD(target_dev);

                    return false;
                }

                Log(string.Format("BLE test Loop {0} OK.", loop_index+1));
            }

            if (exit_test) return false;

            return true;
        }

        private void RunTestBLE()
        {
            bool ret = runTestBLE();
            DelayMs(100);
            BLESerialWrite("scan off\r\n");

            DelayMs(500);
            if (ret)
                Log("RunTestBLE OK.");
            else
                Log("RunTestBLE Fail.");

            //Thread.Sleep(100);
        }

        private void btnBLE_Test_Click(object sender, EventArgs e)
        {
            exit_test = false;
            Task.Factory.StartNew(() => RunTestBLE());
        }

        private void btnLED_Test_Click(object sender, EventArgs e)
        {
            exit_test = false;
            Task.Factory.StartNew(() => RunTestLED());
        }

        private void btn_test_stop_Click(object sender, EventArgs e)
        {
            exit_test = true;
        }

        private void btn_ble_clear_Click(object sender, EventArgs e)
        {

        }

        private void btnBLERefresh_Click(object sender, EventArgs e)
        {

        }

        private void btn_ble_scan_on_Click(object sender, EventArgs e)
        {
            try
            {
                BLE_Recv_Clear();
                BLESerialWrite("scan on\r\n");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private class BLE_Device
        {
            public int RSSI;
            public string UUID;

            public bool Parse(string buff)
            {
                // Device 5A:1D:DF:CA:FB:44  -80

                try
                {
                    string[] ary = buff.Split((char)0x20);

                    if (!ary[1].Contains(":")) return false;
                    UUID = ary[1];
                    RSSI = int.Parse(ary[3]);
                } catch (Exception ex)
                {
                    return false;
                }

                return true;
            }

            public BLE_Device(string buff)
            {
                Parse(buff);
            }
        }
    }
}
