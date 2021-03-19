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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnMotor_Test = new System.Windows.Forms.Button();
            this.btnLED_Test = new System.Windows.Forms.Button();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.chkLED = new System.Windows.Forms.CheckBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.button9 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
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
            this.textBox7 = new System.Windows.Forms.TextBox();
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
            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnMRS1 = new System.Windows.Forms.Button();
            this.txtMotorCount = new System.Windows.Forms.TextBox();
            this.btnMLS1 = new System.Windows.Forms.Button();
            this.txtMotorDelay = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
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
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(495, 546);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "功能測試";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
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
            this.groupBox3.Controls.Add(this.button10);
            this.groupBox3.Controls.Add(this.button11);
            this.groupBox3.Controls.Add(this.textBox8);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.button9);
            this.groupBox3.Controls.Add(this.button8);
            this.groupBox3.Location = new System.Drawing.Point(8, 217);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(480, 121);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "藍芽測試";
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(335, 48);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(75, 23);
            this.button10.TabIndex = 10;
            this.button10.Text = "Disconnect";
            this.button10.UseVisualStyleBackColor = true;
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(254, 48);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(75, 23);
            this.button11.TabIndex = 9;
            this.button11.Text = "Connect";
            this.button11.UseVisualStyleBackColor = true;
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(137, 49);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(111, 22);
            this.textBox8.TabIndex = 8;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(15, 52);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(84, 12);
            this.label11.TabIndex = 7;
            this.label11.Text = "Connect Address";
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(335, 20);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(75, 23);
            this.button9.TabIndex = 6;
            this.button9.Text = "Scan off";
            this.button9.UseVisualStyleBackColor = true;
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(254, 20);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(75, 23);
            this.button8.TabIndex = 5;
            this.button8.Text = "Scan on";
            this.button8.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.txtMotorDelay);
            this.groupBox2.Controls.Add(this.btnMLS1);
            this.groupBox2.Controls.Add(this.txtMotorCount);
            this.groupBox2.Controls.Add(this.btnMRS1);
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
            this.btnSaveSetting.Location = new System.Drawing.Point(416, 305);
            this.btnSaveSetting.Name = "btnSaveSetting";
            this.btnSaveSetting.Size = new System.Drawing.Size(70, 21);
            this.btnSaveSetting.TabIndex = 16;
            this.btnSaveSetting.Text = "Save";
            this.btnSaveSetting.UseVisualStyleBackColor = true;
            this.btnSaveSetting.Click += new System.EventHandler(this.btnSaveSetting_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.textBox7);
            this.groupBox6.Controls.Add(this.label10);
            this.groupBox6.Location = new System.Drawing.Point(6, 178);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(480, 121);
            this.groupBox6.TabIndex = 15;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "藍芽測試";
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(137, 21);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(111, 22);
            this.textBox7.TabIndex = 6;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(15, 24);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(82, 12);
            this.label10.TabIndex = 5;
            this.label10.Text = "Scan RSSI Limit";
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
            this.btnRefresh.Location = new System.Drawing.Point(330, 4);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(36, 36);
            this.btnRefresh.TabIndex = 10;
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // cboPort
            // 
            this.cboPort.FormattingEnabled = true;
            this.cboPort.Location = new System.Drawing.Point(111, 12);
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
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "COM Port";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(254, 12);
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
            this.btnOpen.Location = new System.Drawing.Point(254, 12);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(70, 21);
            this.btnOpen.TabIndex = 11;
            this.btnOpen.Text = "Connect";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // txtLog
            // 
            this.txtLog.BackColor = System.Drawing.Color.Black;
            this.txtLog.Font = new System.Drawing.Font("Consolas", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLog.ForeColor = System.Drawing.Color.White;
            this.txtLog.Location = new System.Drawing.Point(512, 68);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(405, 546);
            this.txtLog.TabIndex = 12;
            // 
            // btnMRS1
            // 
            this.btnMRS1.Location = new System.Drawing.Point(294, 107);
            this.btnMRS1.Name = "btnMRS1";
            this.btnMRS1.Size = new System.Drawing.Size(75, 23);
            this.btnMRS1.TabIndex = 27;
            this.btnMRS1.Text = "MRS";
            this.btnMRS1.UseVisualStyleBackColor = true;
            this.btnMRS1.Click += new System.EventHandler(this.btnMRS1_Click);
            // 
            // txtMotorCount
            // 
            this.txtMotorCount.Location = new System.Drawing.Point(177, 107);
            this.txtMotorCount.Name = "txtMotorCount";
            this.txtMotorCount.Size = new System.Drawing.Size(111, 22);
            this.txtMotorCount.TabIndex = 28;
            this.txtMotorCount.Text = "10";
            // 
            // btnMLS1
            // 
            this.btnMLS1.Location = new System.Drawing.Point(375, 107);
            this.btnMLS1.Name = "btnMLS1";
            this.btnMLS1.Size = new System.Drawing.Size(75, 23);
            this.btnMLS1.TabIndex = 29;
            this.btnMLS1.Text = "MLS";
            this.btnMLS1.UseVisualStyleBackColor = true;
            this.btnMLS1.Click += new System.EventHandler(this.btnMLS1_Click);
            // 
            // txtMotorDelay
            // 
            this.txtMotorDelay.Location = new System.Drawing.Point(29, 109);
            this.txtMotorDelay.Name = "txtMotorDelay";
            this.txtMotorDelay.Size = new System.Drawing.Size(111, 22);
            this.txtMotorDelay.TabIndex = 30;
            this.txtMotorDelay.Text = "100";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(149, 112);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(18, 12);
            this.label3.TabIndex = 11;
            this.label3.Text = "ms";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(921, 620);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.txtLog);
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
        private System.Windows.Forms.TextBox txtLog;
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
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button btnSUV1;
        private System.Windows.Forms.Button btnSUV0;
        private System.Windows.Forms.Button btnWSL1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.TextBox textBox7;
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
        private System.Windows.Forms.TextBox txtMotorDelay;
        private System.Windows.Forms.Button btnMLS1;
        private System.Windows.Forms.TextBox txtMotorCount;
        private System.Windows.Forms.Button btnMRS1;
        private System.Windows.Forms.Label label3;
    }
}

