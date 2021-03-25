namespace Spectrum_Test
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btn_test_stop = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnBLE_Test = new System.Windows.Forms.Button();
            this.btnMotor_Test = new System.Windows.Forms.Button();
            this.btnLED_Test = new System.Windows.Forms.Button();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.chkLED = new System.Windows.Forms.CheckBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btn_ble_write = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btn_ble_send = new System.Windows.Forms.Button();
            this.txt_ble_cmd = new System.Windows.Forms.TextBox();
            this.btn_ble_tab = new System.Windows.Forms.Button();
            this.btn_ble_clear = new System.Windows.Forms.Button();
            this.btn_ble_swl = new System.Windows.Forms.Button();
            this.txt_ble_charact = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_ble_noti = new System.Windows.Forms.Button();
            this.btn_ble_charact = new System.Windows.Forms.Button();
            this.btn_ble_services = new System.Windows.Forms.Button();
            this.txt_ble_service = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_ble_devices = new System.Windows.Forms.Button();
            this.btn_ble_disconnect = new System.Windows.Forms.Button();
            this.btn_ble_connect = new System.Windows.Forms.Button();
            this.txt_ble_uuid = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.btn_ble_scan_off = new System.Windows.Forms.Button();
            this.btn_ble_scan_on = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdoPos2 = new System.Windows.Forms.RadioButton();
            this.rdoPos1 = new System.Windows.Forms.RadioButton();
            this.txtMotorSteps = new System.Windows.Forms.TextBox();
            this.btnMSP = new System.Windows.Forms.Button();
            this.btnDIR1 = new System.Windows.Forms.Button();
            this.btnDIR0 = new System.Windows.Forms.Button();
            this.btnQPS2 = new System.Windows.Forms.Button();
            this.btnQPS1 = new System.Windows.Forms.Button();
            this.btnMLS = new System.Windows.Forms.Button();
            this.btnMRS = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSUV1 = new System.Windows.Forms.Button();
            this.btnSUV0 = new System.Windows.Forms.Button();
            this.btnWSL1 = new System.Windows.Forms.Button();
            this.btnWSL0 = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btnSaveSetting = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.txt_ble_loop_count = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_ble_pass_rssi = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_ble_test_delay = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txt_ble_test_count = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_ble_scan_limit = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.txt_motor_test_round = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.txt_test_led_time = new System.Windows.Forms.TextBox();
            this.chk_test_led3 = new System.Windows.Forms.CheckBox();
            this.chk_test_led2 = new System.Windows.Forms.CheckBox();
            this.chk_test_led1 = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.serialPort = new System.IO.Ports.SerialPort(this.components);
            this.cboBLEComport = new System.Windows.Forms.ComboBox();
            this.btnBLEClose = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnBLERefresh = new System.Windows.Forms.Button();
            this.btnBLEOpen = new System.Windows.Forms.Button();
            this.serialBLEPort = new System.IO.Ports.SerialPort(this.components);
            this.txtLog = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(3, 46);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(503, 572);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btn_test_stop);
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(495, 546);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "功能測試";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btn_test_stop
            // 
            this.btn_test_stop.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_test_stop.Location = new System.Drawing.Point(397, 488);
            this.btn_test_stop.Name = "btn_test_stop";
            this.btn_test_stop.Size = new System.Drawing.Size(84, 41);
            this.btn_test_stop.TabIndex = 14;
            this.btn_test_stop.Text = "Stop";
            this.btn_test_stop.UseVisualStyleBackColor = true;
            this.btn_test_stop.Click += new System.EventHandler(this.btn_test_stop_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnBLE_Test);
            this.panel1.Controls.Add(this.btnMotor_Test);
            this.panel1.Controls.Add(this.btnLED_Test);
            this.panel1.Controls.Add(this.checkBox3);
            this.panel1.Controls.Add(this.checkBox2);
            this.panel1.Controls.Add(this.checkBox1);
            this.panel1.Controls.Add(this.chkLED);
            this.panel1.Location = new System.Drawing.Point(8, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(420, 168);
            this.panel1.TabIndex = 13;
            // 
            // btnBLE_Test
            // 
            this.btnBLE_Test.Location = new System.Drawing.Point(158, 85);
            this.btnBLE_Test.Name = "btnBLE_Test";
            this.btnBLE_Test.Size = new System.Drawing.Size(75, 23);
            this.btnBLE_Test.TabIndex = 8;
            this.btnBLE_Test.Text = "測試";
            this.btnBLE_Test.UseVisualStyleBackColor = true;
            this.btnBLE_Test.Click += new System.EventHandler(this.btnBLE_Test_Click);
            // 
            // btnMotor_Test
            // 
            this.btnMotor_Test.Location = new System.Drawing.Point(158, 49);
            this.btnMotor_Test.Name = "btnMotor_Test";
            this.btnMotor_Test.Size = new System.Drawing.Size(75, 23);
            this.btnMotor_Test.TabIndex = 7;
            this.btnMotor_Test.Text = "測試";
            this.btnMotor_Test.UseVisualStyleBackColor = true;
            this.btnMotor_Test.Click += new System.EventHandler(this.btnMotor_Test_Click);
            // 
            // btnLED_Test
            // 
            this.btnLED_Test.Location = new System.Drawing.Point(158, 15);
            this.btnLED_Test.Name = "btnLED_Test";
            this.btnLED_Test.Size = new System.Drawing.Size(75, 23);
            this.btnLED_Test.TabIndex = 6;
            this.btnLED_Test.Text = "測試";
            this.btnLED_Test.UseVisualStyleBackColor = true;
            this.btnLED_Test.Click += new System.EventHandler(this.btnLED_Test_Click);
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(16, 124);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(96, 16);
            this.checkBox3.TabIndex = 3;
            this.checkBox3.Text = "測試光譜穩定";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(16, 89);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(96, 16);
            this.checkBox2.TabIndex = 2;
            this.checkBox2.Text = "測試藍芽連線";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(16, 53);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(96, 16);
            this.checkBox1.TabIndex = 1;
            this.checkBox1.Text = "測試馬達旋轉";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // chkLED
            // 
            this.chkLED.AutoSize = true;
            this.chkLED.Location = new System.Drawing.Point(16, 19);
            this.chkLED.Name = "chkLED";
            this.chkLED.Size = new System.Drawing.Size(70, 16);
            this.chkLED.TabIndex = 0;
            this.chkLED.Text = "測試LED";
            this.chkLED.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(495, 546);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "基本測試";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btn_ble_write);
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Controls.Add(this.btn_ble_send);
            this.groupBox3.Controls.Add(this.txt_ble_cmd);
            this.groupBox3.Controls.Add(this.btn_ble_tab);
            this.groupBox3.Controls.Add(this.btn_ble_clear);
            this.groupBox3.Controls.Add(this.btn_ble_swl);
            this.groupBox3.Controls.Add(this.txt_ble_charact);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.btn_ble_noti);
            this.groupBox3.Controls.Add(this.btn_ble_charact);
            this.groupBox3.Controls.Add(this.btn_ble_services);
            this.groupBox3.Controls.Add(this.txt_ble_service);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.btn_ble_devices);
            this.groupBox3.Controls.Add(this.btn_ble_disconnect);
            this.groupBox3.Controls.Add(this.btn_ble_connect);
            this.groupBox3.Controls.Add(this.txt_ble_uuid);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.btn_ble_scan_off);
            this.groupBox3.Controls.Add(this.btn_ble_scan_on);
            this.groupBox3.Location = new System.Drawing.Point(8, 217);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(480, 183);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "藍芽測試";
            // 
            // btn_ble_write
            // 
            this.btn_ble_write.Location = new System.Drawing.Point(231, 132);
            this.btn_ble_write.Name = "btn_ble_write";
            this.btn_ble_write.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_write.TabIndex = 25;
            this.btn_ble_write.Text = "Write";
            this.btn_ble_write.UseVisualStyleBackColor = true;
            this.btn_ble_write.Click += new System.EventHandler(this.btn_ble_write_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(13, 133);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(212, 22);
            this.textBox1.TabIndex = 24;
            // 
            // btn_ble_send
            // 
            this.btn_ble_send.Location = new System.Drawing.Point(231, 104);
            this.btn_ble_send.Name = "btn_ble_send";
            this.btn_ble_send.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_send.TabIndex = 23;
            this.btn_ble_send.Text = "Send";
            this.btn_ble_send.UseVisualStyleBackColor = true;
            this.btn_ble_send.Click += new System.EventHandler(this.btn_ble_send_Click);
            // 
            // txt_ble_cmd
            // 
            this.txt_ble_cmd.Location = new System.Drawing.Point(13, 105);
            this.txt_ble_cmd.Name = "txt_ble_cmd";
            this.txt_ble_cmd.Size = new System.Drawing.Size(212, 22);
            this.txt_ble_cmd.TabIndex = 22;
            // 
            // btn_ble_tab
            // 
            this.btn_ble_tab.Location = new System.Drawing.Point(312, 107);
            this.btn_ble_tab.Name = "btn_ble_tab";
            this.btn_ble_tab.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_tab.TabIndex = 21;
            this.btn_ble_tab.Text = "Tab";
            this.btn_ble_tab.UseVisualStyleBackColor = true;
            this.btn_ble_tab.Click += new System.EventHandler(this.btn_ble_tab_Click);
            // 
            // btn_ble_clear
            // 
            this.btn_ble_clear.Location = new System.Drawing.Point(393, 50);
            this.btn_ble_clear.Name = "btn_ble_clear";
            this.btn_ble_clear.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_clear.TabIndex = 20;
            this.btn_ble_clear.Text = "Clear";
            this.btn_ble_clear.UseVisualStyleBackColor = true;
            this.btn_ble_clear.Click += new System.EventHandler(this.btn_ble_clear_Click);
            // 
            // btn_ble_swl
            // 
            this.btn_ble_swl.Location = new System.Drawing.Point(393, 107);
            this.btn_ble_swl.Name = "btn_ble_swl";
            this.btn_ble_swl.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_swl.TabIndex = 19;
            this.btn_ble_swl.Text = "SWL0";
            this.btn_ble_swl.UseVisualStyleBackColor = true;
            this.btn_ble_swl.Click += new System.EventHandler(this.btn_ble_swl_Click);
            // 
            // txt_ble_charact
            // 
            this.txt_ble_charact.Location = new System.Drawing.Point(114, 77);
            this.txt_ble_charact.Name = "txt_ble_charact";
            this.txt_ble_charact.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_charact.TabIndex = 18;
            this.txt_ble_charact.Text = "2";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 80);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 12);
            this.label5.TabIndex = 17;
            this.label5.Text = "Characteristic";
            // 
            // btn_ble_noti
            // 
            this.btn_ble_noti.Location = new System.Drawing.Point(393, 78);
            this.btn_ble_noti.Name = "btn_ble_noti";
            this.btn_ble_noti.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_noti.TabIndex = 16;
            this.btn_ble_noti.Text = "Noti.";
            this.btn_ble_noti.UseVisualStyleBackColor = true;
            this.btn_ble_noti.Click += new System.EventHandler(this.btn_ble_noti_Click);
            // 
            // btn_ble_charact
            // 
            this.btn_ble_charact.Location = new System.Drawing.Point(312, 78);
            this.btn_ble_charact.Name = "btn_ble_charact";
            this.btn_ble_charact.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_charact.TabIndex = 15;
            this.btn_ble_charact.Text = "Charact.";
            this.btn_ble_charact.UseVisualStyleBackColor = true;
            this.btn_ble_charact.Click += new System.EventHandler(this.btn_ble_charact_Click);
            // 
            // btn_ble_services
            // 
            this.btn_ble_services.Location = new System.Drawing.Point(231, 77);
            this.btn_ble_services.Name = "btn_ble_services";
            this.btn_ble_services.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_services.TabIndex = 14;
            this.btn_ble_services.Text = "Services";
            this.btn_ble_services.UseVisualStyleBackColor = true;
            this.btn_ble_services.Click += new System.EventHandler(this.btn_ble_services_Click);
            // 
            // txt_ble_service
            // 
            this.txt_ble_service.Location = new System.Drawing.Point(114, 49);
            this.txt_ble_service.Name = "txt_ble_service";
            this.txt_ble_service.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_service.TabIndex = 13;
            this.txt_ble_service.Text = "1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(11, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(39, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "Service";
            // 
            // btn_ble_devices
            // 
            this.btn_ble_devices.Location = new System.Drawing.Point(393, 21);
            this.btn_ble_devices.Name = "btn_ble_devices";
            this.btn_ble_devices.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_devices.TabIndex = 11;
            this.btn_ble_devices.Text = "Devices";
            this.btn_ble_devices.UseVisualStyleBackColor = true;
            this.btn_ble_devices.Click += new System.EventHandler(this.btn_ble_devices_Click);
            // 
            // btn_ble_disconnect
            // 
            this.btn_ble_disconnect.Location = new System.Drawing.Point(312, 49);
            this.btn_ble_disconnect.Name = "btn_ble_disconnect";
            this.btn_ble_disconnect.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_disconnect.TabIndex = 10;
            this.btn_ble_disconnect.Text = "Disconnect";
            this.btn_ble_disconnect.UseVisualStyleBackColor = true;
            this.btn_ble_disconnect.Click += new System.EventHandler(this.btn_ble_disconnect_Click);
            // 
            // btn_ble_connect
            // 
            this.btn_ble_connect.Location = new System.Drawing.Point(231, 49);
            this.btn_ble_connect.Name = "btn_ble_connect";
            this.btn_ble_connect.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_connect.TabIndex = 9;
            this.btn_ble_connect.Text = "Connect";
            this.btn_ble_connect.UseVisualStyleBackColor = true;
            this.btn_ble_connect.Click += new System.EventHandler(this.btn_ble_connect_Click);
            // 
            // txt_ble_uuid
            // 
            this.txt_ble_uuid.Location = new System.Drawing.Point(114, 23);
            this.txt_ble_uuid.Name = "txt_ble_uuid";
            this.txt_ble_uuid.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_uuid.TabIndex = 8;
            this.txt_ble_uuid.Text = "D8:34:33:F0:08:26";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(11, 26);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(84, 12);
            this.label11.TabIndex = 7;
            this.label11.Text = "Connect Address";
            // 
            // btn_ble_scan_off
            // 
            this.btn_ble_scan_off.Location = new System.Drawing.Point(312, 21);
            this.btn_ble_scan_off.Name = "btn_ble_scan_off";
            this.btn_ble_scan_off.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_scan_off.TabIndex = 6;
            this.btn_ble_scan_off.Text = "Scan off";
            this.btn_ble_scan_off.UseVisualStyleBackColor = true;
            this.btn_ble_scan_off.Click += new System.EventHandler(this.btn_ble_scan_off_Click);
            // 
            // btn_ble_scan_on
            // 
            this.btn_ble_scan_on.Location = new System.Drawing.Point(231, 21);
            this.btn_ble_scan_on.Name = "btn_ble_scan_on";
            this.btn_ble_scan_on.Size = new System.Drawing.Size(75, 23);
            this.btn_ble_scan_on.TabIndex = 5;
            this.btn_ble_scan_on.Text = "Scan on";
            this.btn_ble_scan_on.UseVisualStyleBackColor = true;
            this.btn_ble_scan_on.Click += new System.EventHandler(this.btn_ble_scan_on_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdoPos2);
            this.groupBox2.Controls.Add(this.rdoPos1);
            this.groupBox2.Controls.Add(this.txtMotorSteps);
            this.groupBox2.Controls.Add(this.btnMSP);
            this.groupBox2.Controls.Add(this.btnDIR1);
            this.groupBox2.Controls.Add(this.btnDIR0);
            this.groupBox2.Controls.Add(this.btnQPS2);
            this.groupBox2.Controls.Add(this.btnQPS1);
            this.groupBox2.Controls.Add(this.btnMLS);
            this.groupBox2.Controls.Add(this.btnMRS);
            this.groupBox2.Location = new System.Drawing.Point(8, 68);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(480, 143);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "馬達旋轉測試";
            // 
            // rdoPos2
            // 
            this.rdoPos2.AutoSize = true;
            this.rdoPos2.Location = new System.Drawing.Point(146, 53);
            this.rdoPos2.Name = "rdoPos2";
            this.rdoPos2.Size = new System.Drawing.Size(45, 16);
            this.rdoPos2.TabIndex = 26;
            this.rdoPos2.Text = "Pos2";
            this.rdoPos2.UseVisualStyleBackColor = true;
            // 
            // rdoPos1
            // 
            this.rdoPos1.AutoSize = true;
            this.rdoPos1.Checked = true;
            this.rdoPos1.Location = new System.Drawing.Point(95, 53);
            this.rdoPos1.Name = "rdoPos1";
            this.rdoPos1.Size = new System.Drawing.Size(45, 16);
            this.rdoPos1.TabIndex = 25;
            this.rdoPos1.TabStop = true;
            this.rdoPos1.Text = "Pos1";
            this.rdoPos1.UseVisualStyleBackColor = true;
            // 
            // txtMotorSteps
            // 
            this.txtMotorSteps.Location = new System.Drawing.Point(177, 79);
            this.txtMotorSteps.Name = "txtMotorSteps";
            this.txtMotorSteps.Size = new System.Drawing.Size(111, 22);
            this.txtMotorSteps.TabIndex = 24;
            this.txtMotorSteps.Text = "500";
            // 
            // btnMSP
            // 
            this.btnMSP.Location = new System.Drawing.Point(14, 50);
            this.btnMSP.Name = "btnMSP";
            this.btnMSP.Size = new System.Drawing.Size(75, 23);
            this.btnMSP.TabIndex = 21;
            this.btnMSP.Text = "MSP";
            this.btnMSP.UseVisualStyleBackColor = true;
            this.btnMSP.Click += new System.EventHandler(this.btnMSP_Click);
            // 
            // btnDIR1
            // 
            this.btnDIR1.Location = new System.Drawing.Point(96, 21);
            this.btnDIR1.Name = "btnDIR1";
            this.btnDIR1.Size = new System.Drawing.Size(75, 23);
            this.btnDIR1.TabIndex = 20;
            this.btnDIR1.Text = "DIR1";
            this.btnDIR1.UseVisualStyleBackColor = true;
            this.btnDIR1.Click += new System.EventHandler(this.btnDIR1_Click);
            // 
            // btnDIR0
            // 
            this.btnDIR0.Location = new System.Drawing.Point(14, 21);
            this.btnDIR0.Name = "btnDIR0";
            this.btnDIR0.Size = new System.Drawing.Size(75, 23);
            this.btnDIR0.TabIndex = 19;
            this.btnDIR0.Text = "DIR0";
            this.btnDIR0.UseVisualStyleBackColor = true;
            this.btnDIR0.Click += new System.EventHandler(this.btnDIR0_Click);
            // 
            // btnQPS2
            // 
            this.btnQPS2.Location = new System.Drawing.Point(258, 21);
            this.btnQPS2.Name = "btnQPS2";
            this.btnQPS2.Size = new System.Drawing.Size(75, 23);
            this.btnQPS2.TabIndex = 18;
            this.btnQPS2.Text = "QPS2";
            this.btnQPS2.UseVisualStyleBackColor = true;
            this.btnQPS2.Click += new System.EventHandler(this.btnQPS2_Click);
            // 
            // btnQPS1
            // 
            this.btnQPS1.Location = new System.Drawing.Point(177, 21);
            this.btnQPS1.Name = "btnQPS1";
            this.btnQPS1.Size = new System.Drawing.Size(75, 23);
            this.btnQPS1.TabIndex = 17;
            this.btnQPS1.Text = "QPS1";
            this.btnQPS1.UseVisualStyleBackColor = true;
            this.btnQPS1.Click += new System.EventHandler(this.btnQPS1_Click);
            // 
            // btnMLS
            // 
            this.btnMLS.Location = new System.Drawing.Point(95, 79);
            this.btnMLS.Name = "btnMLS";
            this.btnMLS.Size = new System.Drawing.Size(75, 23);
            this.btnMLS.TabIndex = 15;
            this.btnMLS.Text = "MLS";
            this.btnMLS.UseVisualStyleBackColor = true;
            this.btnMLS.Click += new System.EventHandler(this.btnMLS_Click);
            // 
            // btnMRS
            // 
            this.btnMRS.Location = new System.Drawing.Point(14, 79);
            this.btnMRS.Name = "btnMRS";
            this.btnMRS.Size = new System.Drawing.Size(75, 23);
            this.btnMRS.TabIndex = 7;
            this.btnMRS.Text = "MRS";
            this.btnMRS.UseVisualStyleBackColor = true;
            this.btnMRS.Click += new System.EventHandler(this.btnMRS_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSUV1);
            this.groupBox1.Controls.Add(this.btnSUV0);
            this.groupBox1.Controls.Add(this.btnWSL1);
            this.groupBox1.Controls.Add(this.btnWSL0);
            this.groupBox1.Location = new System.Drawing.Point(8, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(480, 56);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "LED測試";
            // 
            // btnSUV1
            // 
            this.btnSUV1.Location = new System.Drawing.Point(258, 21);
            this.btnSUV1.Name = "btnSUV1";
            this.btnSUV1.Size = new System.Drawing.Size(75, 23);
            this.btnSUV1.TabIndex = 8;
            this.btnSUV1.Text = "SUV1";
            this.btnSUV1.UseVisualStyleBackColor = true;
            this.btnSUV1.Click += new System.EventHandler(this.btnSUV1_Click);
            // 
            // btnSUV0
            // 
            this.btnSUV0.Location = new System.Drawing.Point(177, 21);
            this.btnSUV0.Name = "btnSUV0";
            this.btnSUV0.Size = new System.Drawing.Size(75, 23);
            this.btnSUV0.TabIndex = 7;
            this.btnSUV0.Text = "SUV0";
            this.btnSUV0.UseVisualStyleBackColor = true;
            this.btnSUV0.Click += new System.EventHandler(this.btnSUV0_Click);
            // 
            // btnWSL1
            // 
            this.btnWSL1.Location = new System.Drawing.Point(96, 21);
            this.btnWSL1.Name = "btnWSL1";
            this.btnWSL1.Size = new System.Drawing.Size(75, 23);
            this.btnWSL1.TabIndex = 6;
            this.btnWSL1.Text = "SWL1";
            this.btnWSL1.UseVisualStyleBackColor = true;
            this.btnWSL1.Click += new System.EventHandler(this.btnWSL1_Click);
            // 
            // btnWSL0
            // 
            this.btnWSL0.Location = new System.Drawing.Point(14, 21);
            this.btnWSL0.Name = "btnWSL0";
            this.btnWSL0.Size = new System.Drawing.Size(75, 23);
            this.btnWSL0.TabIndex = 5;
            this.btnWSL0.Text = "SWL0";
            this.btnWSL0.UseVisualStyleBackColor = true;
            this.btnWSL0.Click += new System.EventHandler(this.btnWSL0_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(495, 546);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "進階測試";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.btnSaveSetting);
            this.tabPage4.Controls.Add(this.groupBox6);
            this.tabPage4.Controls.Add(this.groupBox5);
            this.tabPage4.Controls.Add(this.groupBox4);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(495, 546);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "測試設定";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btnSaveSetting
            // 
            this.btnSaveSetting.Location = new System.Drawing.Point(416, 360);
            this.btnSaveSetting.Name = "btnSaveSetting";
            this.btnSaveSetting.Size = new System.Drawing.Size(70, 21);
            this.btnSaveSetting.TabIndex = 16;
            this.btnSaveSetting.Text = "Save";
            this.btnSaveSetting.UseVisualStyleBackColor = true;
            this.btnSaveSetting.Click += new System.EventHandler(this.btnSaveSetting_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.txt_ble_loop_count);
            this.groupBox6.Controls.Add(this.label9);
            this.groupBox6.Controls.Add(this.txt_ble_pass_rssi);
            this.groupBox6.Controls.Add(this.label8);
            this.groupBox6.Controls.Add(this.txt_ble_test_delay);
            this.groupBox6.Controls.Add(this.label7);
            this.groupBox6.Controls.Add(this.txt_ble_test_count);
            this.groupBox6.Controls.Add(this.label6);
            this.groupBox6.Controls.Add(this.txt_ble_scan_limit);
            this.groupBox6.Controls.Add(this.label10);
            this.groupBox6.Location = new System.Drawing.Point(6, 178);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(480, 164);
            this.groupBox6.TabIndex = 15;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "藍芽測試";
            // 
            // txt_ble_loop_count
            // 
            this.txt_ble_loop_count.Location = new System.Drawing.Point(137, 133);
            this.txt_ble_loop_count.Name = "txt_ble_loop_count";
            this.txt_ble_loop_count.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_loop_count.TabIndex = 14;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 136);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(82, 12);
            this.label9.TabIndex = 13;
            this.label9.Text = "測試LOOP次數";
            // 
            // txt_ble_pass_rssi
            // 
            this.txt_ble_pass_rssi.Location = new System.Drawing.Point(137, 49);
            this.txt_ble_pass_rssi.Name = "txt_ble_pass_rssi";
            this.txt_ble_pass_rssi.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_pass_rssi.TabIndex = 12;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(15, 52);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(56, 12);
            this.label8.TabIndex = 11;
            this.label8.Text = "合格 RSSI";
            // 
            // txt_ble_test_delay
            // 
            this.txt_ble_test_delay.Location = new System.Drawing.Point(137, 105);
            this.txt_ble_test_delay.Name = "txt_ble_test_delay";
            this.txt_ble_test_delay.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_test_delay.TabIndex = 10;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 108);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 12);
            this.label7.TabIndex = 9;
            this.label7.Text = "指令延遲秒數";
            // 
            // txt_ble_test_count
            // 
            this.txt_ble_test_count.Location = new System.Drawing.Point(137, 77);
            this.txt_ble_test_count.Name = "txt_ble_test_count";
            this.txt_ble_test_count.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_test_count.TabIndex = 8;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 80);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 12);
            this.label6.TabIndex = 7;
            this.label6.Text = "指令測試次數";
            // 
            // txt_ble_scan_limit
            // 
            this.txt_ble_scan_limit.Location = new System.Drawing.Point(137, 21);
            this.txt_ble_scan_limit.Name = "txt_ble_scan_limit";
            this.txt_ble_scan_limit.Size = new System.Drawing.Size(111, 22);
            this.txt_ble_scan_limit.TabIndex = 6;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(15, 24);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(84, 12);
            this.label10.TabIndex = 5;
            this.label10.Text = "掃描 RSSI Limit";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.txt_motor_test_round);
            this.groupBox5.Controls.Add(this.label2);
            this.groupBox5.Controls.Add(this.label16);
            this.groupBox5.Controls.Add(this.label17);
            this.groupBox5.Controls.Add(this.textBox10);
            this.groupBox5.Location = new System.Drawing.Point(6, 87);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(480, 85);
            this.groupBox5.TabIndex = 2;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "馬達設定";
            // 
            // txt_motor_test_round
            // 
            this.txt_motor_test_round.Location = new System.Drawing.Point(137, 49);
            this.txt_motor_test_round.Name = "txt_motor_test_round";
            this.txt_motor_test_round.Size = new System.Drawing.Size(111, 22);
            this.txt_motor_test_round.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 22;
            this.label2.Text = "馬達測試來回趟數";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(256, 24);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(17, 12);
            this.label16.TabIndex = 11;
            this.label16.Text = "步";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(12, 24);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(119, 12);
            this.label17.TabIndex = 9;
            this.label17.Text = "步進馬達定義  1mm = ";
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(137, 21);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(111, 22);
            this.textBox10.TabIndex = 8;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.txt_test_led_time);
            this.groupBox4.Controls.Add(this.chk_test_led3);
            this.groupBox4.Controls.Add(this.chk_test_led2);
            this.groupBox4.Controls.Add(this.chk_test_led1);
            this.groupBox4.Controls.Add(this.label12);
            this.groupBox4.Location = new System.Drawing.Point(6, 6);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(480, 75);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "LED設定";
            // 
            // txt_test_led_time
            // 
            this.txt_test_led_time.Location = new System.Drawing.Point(99, 43);
            this.txt_test_led_time.Name = "txt_test_led_time";
            this.txt_test_led_time.Size = new System.Drawing.Size(111, 22);
            this.txt_test_led_time.TabIndex = 4;
            // 
            // chk_test_led3
            // 
            this.chk_test_led3.AutoSize = true;
            this.chk_test_led3.Location = new System.Drawing.Point(187, 21);
            this.chk_test_led3.Name = "chk_test_led3";
            this.chk_test_led3.Size = new System.Drawing.Size(79, 16);
            this.chk_test_led3.TabIndex = 3;
            this.chk_test_led3.Text = "測試LED 3";
            this.chk_test_led3.UseVisualStyleBackColor = true;
            this.chk_test_led3.Visible = false;
            // 
            // chk_test_led2
            // 
            this.chk_test_led2.AutoSize = true;
            this.chk_test_led2.Location = new System.Drawing.Point(102, 21);
            this.chk_test_led2.Name = "chk_test_led2";
            this.chk_test_led2.Size = new System.Drawing.Size(79, 16);
            this.chk_test_led2.TabIndex = 2;
            this.chk_test_led2.Text = "測試LED 2";
            this.chk_test_led2.UseVisualStyleBackColor = true;
            // 
            // chk_test_led1
            // 
            this.chk_test_led1.AutoSize = true;
            this.chk_test_led1.Location = new System.Drawing.Point(17, 21);
            this.chk_test_led1.Name = "chk_test_led1";
            this.chk_test_led1.Size = new System.Drawing.Size(79, 16);
            this.chk_test_led1.TabIndex = 1;
            this.chk_test_led1.Text = "測試LED 1";
            this.chk_test_led1.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(15, 46);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(53, 12);
            this.label12.TabIndex = 0;
            this.label12.Text = "點亮秒數";
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackgroundImage = global::Spectrum_Test.Properties.Resources.refresh;
            this.btnRefresh.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnRefresh.Location = new System.Drawing.Point(348, 4);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(36, 36);
            this.btnRefresh.TabIndex = 10;
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // cboPort
            // 
            this.cboPort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPort.FormattingEnabled = true;
            this.cboPort.Location = new System.Drawing.Point(129, 12);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(137, 20);
            this.cboPort.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 15);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "Spectrum COM Port";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(272, 12);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(70, 21);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "Disconnect";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Visible = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(272, 13);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(70, 21);
            this.btnOpen.TabIndex = 11;
            this.btnOpen.Text = "Connect";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // cboBLEComport
            // 
            this.cboBLEComport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboBLEComport.FormattingEnabled = true;
            this.cboBLEComport.Location = new System.Drawing.Point(503, 13);
            this.cboBLEComport.Name = "cboBLEComport";
            this.cboBLEComport.Size = new System.Drawing.Size(137, 20);
            this.cboBLEComport.TabIndex = 14;
            // 
            // btnBLEClose
            // 
            this.btnBLEClose.Location = new System.Drawing.Point(646, 13);
            this.btnBLEClose.Name = "btnBLEClose";
            this.btnBLEClose.Size = new System.Drawing.Size(70, 21);
            this.btnBLEClose.TabIndex = 15;
            this.btnBLEClose.Text = "Disconnect";
            this.btnBLEClose.UseVisualStyleBackColor = true;
            this.btnBLEClose.Visible = false;
            this.btnBLEClose.Click += new System.EventHandler(this.btnBLEClose_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(411, 16);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 12);
            this.label3.TabIndex = 13;
            this.label3.Text = "BLE COM Port";
            // 
            // btnBLERefresh
            // 
            this.btnBLERefresh.BackgroundImage = global::Spectrum_Test.Properties.Resources.refresh;
            this.btnBLERefresh.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnBLERefresh.Location = new System.Drawing.Point(722, 5);
            this.btnBLERefresh.Name = "btnBLERefresh";
            this.btnBLERefresh.Size = new System.Drawing.Size(36, 36);
            this.btnBLERefresh.TabIndex = 16;
            this.btnBLERefresh.UseVisualStyleBackColor = true;
            this.btnBLERefresh.Click += new System.EventHandler(this.btnBLERefresh_Click);
            // 
            // btnBLEOpen
            // 
            this.btnBLEOpen.Location = new System.Drawing.Point(646, 40);
            this.btnBLEOpen.Name = "btnBLEOpen";
            this.btnBLEOpen.Size = new System.Drawing.Size(70, 21);
            this.btnBLEOpen.TabIndex = 17;
            this.btnBLEOpen.Text = "Connect";
            this.btnBLEOpen.UseVisualStyleBackColor = true;
            this.btnBLEOpen.Click += new System.EventHandler(this.btnBLEOpen_Click);
            // 
            // serialBLEPort
            // 
            this.serialBLEPort.ReadTimeout = 10;
            // 
            // txtLog
            // 
            this.txtLog.BackColor = System.Drawing.Color.Black;
            this.txtLog.Font = new System.Drawing.Font("Consolas", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLog.ForeColor = System.Drawing.Color.White;
            this.txtLog.Location = new System.Drawing.Point(508, 68);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(410, 546);
            this.txtLog.TabIndex = 18;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(921, 620);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnBLEOpen);
            this.Controls.Add(this.cboBLEComport);
            this.Controls.Add(this.btnBLEClose);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnBLERefresh);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.cboPort);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRefresh);
            this.Name = "MainForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.ComboBox cboPort;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnClose;
        private System.IO.Ports.SerialPort serialPort;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox chkLED;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnMSP;
        private System.Windows.Forms.Button btnDIR1;
        private System.Windows.Forms.Button btnDIR0;
        private System.Windows.Forms.Button btnQPS2;
        private System.Windows.Forms.Button btnQPS1;
        private System.Windows.Forms.Button btnMLS;
        private System.Windows.Forms.Button btnMRS;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnWSL0;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btn_ble_disconnect;
        private System.Windows.Forms.Button btn_ble_connect;
        private System.Windows.Forms.TextBox txt_ble_uuid;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btn_ble_scan_off;
        private System.Windows.Forms.Button btn_ble_scan_on;
        private System.Windows.Forms.Button btnSUV1;
        private System.Windows.Forms.Button btnSUV0;
        private System.Windows.Forms.Button btnWSL1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.TextBox txt_ble_scan_limit;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.TextBox txt_motor_test_round;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox txt_test_led_time;
        private System.Windows.Forms.CheckBox chk_test_led3;
        private System.Windows.Forms.CheckBox chk_test_led2;
        private System.Windows.Forms.CheckBox chk_test_led1;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button btnLED_Test;
        private System.Windows.Forms.TextBox txtMotorSteps;
        private System.Windows.Forms.Button btnSaveSetting;
        private System.Windows.Forms.RadioButton rdoPos2;
        private System.Windows.Forms.RadioButton rdoPos1;
        private System.Windows.Forms.Button btnMotor_Test;
        private System.Windows.Forms.ComboBox cboBLEComport;
        private System.Windows.Forms.Button btnBLEClose;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnBLERefresh;
        private System.Windows.Forms.Button btnBLEOpen;
        private System.Windows.Forms.Button btnBLE_Test;
        private System.IO.Ports.SerialPort serialBLEPort;
        private System.Windows.Forms.Button btn_ble_devices;
        private System.Windows.Forms.TextBox txt_ble_service;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btn_ble_swl;
        private System.Windows.Forms.TextBox txt_ble_charact;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btn_ble_noti;
        private System.Windows.Forms.Button btn_ble_charact;
        private System.Windows.Forms.Button btn_ble_services;
        private System.Windows.Forms.Button btn_ble_send;
        private System.Windows.Forms.TextBox txt_ble_cmd;
        private System.Windows.Forms.Button btn_ble_tab;
        private System.Windows.Forms.Button btn_ble_clear;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Button btn_ble_write;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox txt_ble_test_count;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_ble_pass_rssi;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_ble_test_delay;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt_ble_loop_count;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btn_test_stop;
    }
}

