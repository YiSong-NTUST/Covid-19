using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using GitHub.secile.Video;

namespace SpertroApp
{
    public partial class Form_ini : Form
    {
        public int Gamma = 190;
        public int AG = 2;
        public Form_ini()
        {
            InitializeComponent();
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            Form1 F1 = new Form1(Gamma, AG);
           this.Visible=false;
            F1.Show();
        
        }

        private void btn_Camp_Click(object sender, EventArgs e)
        {
            Form_camp FP = new Form_camp(Gamma, AG);
            this.Visible = false;
            FP.Show();
            Console_Connect(comboBox1.Text, Convert.ToInt32(TB_Baud.Text));

            SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 1;CHAN 2;LAS:OUT 1;CHAN 3;LAS:OUT 1;CHAN 4;LAS:OUT 1;CHAN 5;LAS:OUT 1;CHAN 6;LAS:OUT 1;CHAN 7;LAS:OUT 1;CHAN 8;LAS:OUT 1;\r\n"));
            Thread.Sleep(1000);
            Console_Connect(comboBox1.Text, Convert.ToInt32(TB_Baud.Text));
        }

        //序列
        private SerialPort My_SerialPort;
        public void SendData(Object sendBuffer)
        {
            if (My_SerialPort.IsOpen == false)
                return;
            if (sendBuffer != null)
            {
                Byte[] buffer = sendBuffer as Byte[];

                try
                {
                    My_SerialPort.Write(buffer, 0, buffer.Length);
                }
                catch (Exception ex)
                {
                    CloseComport();
                    MessageBox.Show(ex.Message);
                }
                buffer = null;
            }

        }

        public void CloseComport()
        {
            try
            {
                My_SerialPort.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Console_Connect(string COM, Int32 baud)
        {
            try
            {
                My_SerialPort = new SerialPort();

                if (My_SerialPort.IsOpen)
                {
                    My_SerialPort.Close();
                }

                //設定 Serial Port 參數
                My_SerialPort.PortName = COM;
                My_SerialPort.BaudRate = baud;
                My_SerialPort.DataBits = 8;
                My_SerialPort.StopBits = StopBits.One;
                //波特率
                My_SerialPort.BaudRate = 9600;
                //資料位
                My_SerialPort.DataBits = 8;
                //  My_SerialPort.PortName = Form1.comboBox1.Text;
                //一個停止位
                My_SerialPort.StopBits = System.IO.Ports.StopBits.One;
                //無奇偶校驗位
                My_SerialPort.Parity = System.IO.Ports.Parity.None;
                My_SerialPort.WriteBufferSize = 8192;
                My_SerialPort.ReadTimeout = 1000;

                if (!My_SerialPort.IsOpen)
                {
                    //開啟 Serial Port
                    My_SerialPort.Open();

                    //Console_receiving = true;

                    //開啟執行續做接收動作
                    //t = new Thread(DoReceive);

                    //t.IsBackground = true;
                    //t.Start();
                /*    Console_Input.Text = "";
                    Console_Input.Text = "連結成功";*/
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message);
            }
        }

        private void Form_ini_Load(object sender, EventArgs e)
        {
            string[] devices = UsbCamera.FindDevices();
            comboBox2.Items.Add(devices[0]) ;

            string[] portnames = SerialPort.GetPortNames();
            foreach (var item in portnames)
            {
                comboBox1.Items.Add(item);
            }


        }
        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
