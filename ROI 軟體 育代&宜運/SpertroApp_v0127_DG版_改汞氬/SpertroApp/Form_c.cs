using GitHub.secile.Video;
using Newtonsoft.Json;
using SpectroChipApp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
//視窗截圖
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Windows.Forms.DataVisualization.Charting;
using System.Net;
using System.Collections.Specialized;

namespace SpertroApp
{

    public partial class Form1 : Form
    {
        //AR量測
        public List<double> Ar_Intensity = new List<double>();
        public int gamma_number = 100;
        private bool isGet_Ar = false;
        public List<double> Poly_Coefs_of_Hg_Ar = new List<double>();
        //白光量測Noise
        public double White_Imax_Divided_By_DarkNoise = 0;
        public double White_Imax_Divided_By_Dark = 0;
        //白光AutoScaling
        private bool isAutoScaling_FisrtStep_Complete = false;
        public int back_number =0;
        public bool isGetROI = false;
        delegate void Dg_Update(int dg);
        delegate void Gamma_Update(int gamma);
        delegate void Back_Update(int back);
        List<double> result_buffer = new List<double>();
        List<double> dg_set = new List<double>();
        List<double> Max_Intensity = new List<double>();
        public string final_dg = "";
        public string final_gamma = "";
        public string final_back = "";
        public double Index_of_Max = 0;

        //取平均秀出新光譜
        public bool isAutoScaling_END = false;
        public bool isAvgEnd = false;
        //取DARK圖
        public bool isGetDarkChart = false;
        public bool isGetDarkData = false;
        //Baseline_mid_point
        public double Baseline_wavelength_of_minPoint = 0;
        public double Baseline_of_leftPoint = 0;
        public double Baseline_of_rightPoint = 0;
        //HG 只算2.3 PEAK的RMS
        public bool Hg1WFHM = false;
        public bool Hg2WFHM = false;
        public bool Hg3WFHM = false;
        public bool Hg4WFHM = false;

        //單根模式的疊圖截圖
        public Bitmap p1;
        public Bitmap p2;
        public Bitmap chart_p1;
        public Bitmap chart_p2;
        public bool isOverlay_First = false;

        //單根雷射
        public List<double> Laser_Intensity = new List<double>();
        public List<double> Save_Laser_Intensity = new List<double>();

        //λcal-DIFF-DIFFrms
        public List<double> Lamda_cal = new List<double>();
        public List<double> DIFF = new List<double>();
        public double DIFF_rms = new double();
        //畫波長對強度圖
        public bool isend = false;
        //白光量測
        private bool onlyWhite = false;
        //控件截圖
        delegate void chartScreenShot(string name);
        public bool isReady = false;
        //SCID
        public string this_SC_ID = "";
        //汞氬燈-------------------------------------
        public bool isHg_Ar = false;
        private int isFirtCacular = 0;
        public string report = "";
        public string report_Memory = "";

        //-------------------------------------------
        public bool isLaser = false;
        public List<int> RemoveIndex = new List<int>();
        private bool issave = false;
        public List<double> deltaL = new List<double>();
        public List<double> Wavelength = new List<double>();
        public List<double> Wavelength_save = new List<double>();
        public List<double> pixel_M = new List<double>();
        public List<double> pixel_M_save = new List<double>();
        public List<double> LambdaSC = new List<double>();
        public List<double> DeltaPsc = new List<double>();
        public List<double> LambdaSC_save = new List<double>();
        public List<double> DeltaPsc_save = new List<double>();
        public double LambdaSC_RMS = 0;
        public double LambdaSC_Spot_RMS = 0;
        public double LambdaSC_STD = 0;
        public double LambdaSC_Spot_STD = 0;
        public double LambdaSC_STD_100 = 0;
        public double LambdaSC_Spot_STD_100 = 0;
        public double SNR = 0;
        public double Stray_light = 0;
        public double Dynamic_Range = 0;
        public double max_intensity = 0;
        public double[] Noise = new double[100];
        public double RMS_of_Noise = 0;
        public double Pixel_of_maxIntensity = 0;
        public double lamda_max_Intensity = 0;

        private bool isch1 = false, isch2 = false, isch3 = false, isch4 = false, isch5 = false, isch6 = false, isch7 = false, isch8 = false;
        #region auto_change_form_size
        // <summary>

        /// 窗体改变前的Width宽度

        /// </summary>

        private int iForm_ResizeBefore_Width;

        /// <summary>

        /// 窗体改变前的Height高度

        /// </summary>

        private int iForm_ResizeBefore_Height;



        /// <summary>

        /// 取得Form中控件中初始的字体大小，即无论Form多大此字体都是Form中最小值

        /// </summary>

        private float FontSizeMin = 0; //2020-03-23 19:21:50 更新



        /// <summary>

        /// 宽比例，float类型，此处为每次改后窗体的宽度 除以改变前窗体的宽度的比例

        /// </summary>

        private float fRatio_Width;

        /// <summary>

        /// 高比例，float类型，此处为每次改后窗体的高度 除以改变前窗体的高度的比例

        /// </summary>

        private float fRatio_Height;
        #endregion
        //Camera
        static Size Read_Size = new Size(1280, 960);
        UsbCamera camera = new UsbCamera(0, Read_Size);

        //Live Stop
        public bool bCameraLive = false;

        //ImageBuff
        //private byte[] GrayBuffer;

        //ROI
        Point ROI_Point;
        public int ROI_X = 0;
        public int ROI_Y = 0;
        public int ROI_W = 1280;
        public int ROI_H = 20;
        const int ROI_Edge = 10;

        //console 參數
        private SerialPort My_SerialPort;
        private bool Console_receiving = false;
        private Thread t;
        private int index = 0;
        //使用委派顯示 Console 畫面
        delegate void Display(string buffer);

        private object BufferLock = new object();

        //儲存Intensity
        private List<double> Dark_Intensity = new List<double>();
        private List<double> Original_Intensity = new List<double>();
        private List<double> RealTime_Original_Intensity = new List<double>();
        private List<double> Pure_Intensity = new List<double>();
        private List<double> SG_Intensity = new List<double>();
        // private List<double> Gaus_Intensity = new List<double>();


        //Step1新增參數
        private bool isCalibratingEXP = false;

        private static double scaleX = 1;
        private static double scaleY = 1;

        //校正參數顯示
        private static string Context = "λ\tPixel\tΔP(FWHM)\tΔλm(FWHM)\r\n";

        //校正結果顯示
        private static string Result = "{   }";
        private static string Readble = "△相機參數:\n  曝光(ELC)= 未知ELC\n  數位增益(GNV) = 未知GNV\n  類比增益(AGN) = 未知AGN\n" +
           "△興趣區間:\n  X = 未知X\n  W = 未知\n  Y = 未知\n  H = 未知\n" +
            "△校正係數:\n  a0 = 未知a0\n  a1 = 未知a1\n  a2 = 未知\n  a3 = 未知a3\n  a4 = 未知a4\n";


        private bool Pause = false; //控制Step的參數

        //相機參數
        private Dictionary<string, int> Camera_Parameters = new Dictionary<string, int>();
        //一次八根

        private List<double> FWHL = new List<double>();


        private Bitmap CopyBitmap(byte[] Buffer, int width, int height)
        {

            var result = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            if (Buffer == null) return result;

            var bmp_data = result.LockBits(new Rectangle(Point.Empty, result.Size), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            lock (BufferLock)
            {
                // copy from last row.
                for (int y = 0; y < height; y++)
                {
                    int stride = width * 3;

                    var src_idx = Buffer.Length - (stride * (y + 1));
                    var dst = IntPtr.Add(bmp_data.Scan0, stride * y);

                    Marshal.Copy(Buffer, src_idx, dst, stride);

                }
            }
            result.UnlockBits(bmp_data);

            return result;
        }

        //convert image to bytearray
        public static byte[] ImageToBuffer(Image Image, System.Drawing.Imaging.ImageFormat imageFormat)
        {
            if (Image == null) { return null; }
            byte[] data = null;
            using (MemoryStream oMemoryStream = new MemoryStream())
            {
                //建立副本
                using (Bitmap oBitmap = new Bitmap(Image))
                {
                    //儲存圖片到 MemoryStream 物件，並且指定儲存影像之格式
                    oBitmap.Save(oMemoryStream, imageFormat);
                    //設定資料流位置
                    oMemoryStream.Position = 0;
                    //設定 buffer 長度
                    data = new byte[oMemoryStream.Length];
                    //將資料寫入 buffer
                    oMemoryStream.Read(data, 0, Convert.ToInt32(oMemoryStream.Length));
                    //將所有緩衝區的資料寫入資料流
                    oMemoryStream.Flush();
                }
            }
            return data;
        }


        public static Bitmap BufferToImage(byte[] Buffer) //改
        {
            if (Buffer == null || Buffer.Length == 0) { return null; }
            byte[] data = null;
            Image oImage = null;
            Bitmap oBitmap = null;
            //建立副本
            data = (byte[])Buffer.Clone();
            try
            {
                MemoryStream oMemoryStream = new MemoryStream(Buffer);
                //設定資料流位置
                oMemoryStream.Position = 0;
                oImage = System.Drawing.Image.FromStream(oMemoryStream);
                //建立副本
                oBitmap = new Bitmap(oImage);
            }
            catch
            {
                throw;
            }
            //return oImage;
            return oBitmap;
        }

        delegate void FormUpdata(int roi_y, int proc, string Context, string Result, string Readble);
        void FormUpdataMethod(int roi_y, int proc, string Context, string Result, string Readble)
        {
            // Text = i.ToString();
            DrawCanvas.Top = roi_y;
            // 

            //progressROI.Value = proc;
            richTextBox1.Text = Context;
            richTextBox2.Text = Result;
            richTextBox3.Text = Readble;

        }

        delegate void set_camera_prop_dele(string item, int prop_num);

        // Text = i.ToString();

        private List<int> exp_recorder = new List<int>();
        void set_camera_prop(string item, int prop_num)
        {
            UsbCamera.PropertyItems.Property prop;
            switch (item)
            {
                case "exp":
                    prop = camera.Properties[DirectShow.CameraControlProperty.Exposure];
                    if (prop.Available && prop.CanAuto)
                    {
                        prop.SetValue(DirectShow.CameraControlFlags.Auto, 0);
                    }
                    break;

                case "bright":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Brightness];
                    break;
                case "con":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Contrast];
                    break;
                case "hue":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Hue];
                    break;
                case "sat":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Saturation];
                    break;
                case "sharp":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Sharpness];
                    break;
                case "gamma":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gamma];
                    break;
                case "white":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.WhiteBalance];
                    break;
                case "back":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.BacklightCompensation];
                    break;
                case "gain":
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
                    break;
                default:
                    prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
                    break;


            }

            if (prop.Available)
            {
                var min = prop.Min;
                var max = prop.Max;
                var def = prop.Default;
                var step = prop.Step;


                if (prop_num <= min)
                {
                    prop.Default = min;
                }
                else if (prop_num >= max)
                {

                    prop.Default = max;
                }
                else
                {
                    prop.Default = prop_num;
                }
                exp_recorder.Add(prop.Default);
                Console.WriteLine(exp_recorder);
                prop.SetValue(DirectShow.CameraControlFlags.Manual, prop.Default);




            }

            var q = from p in exp_recorder
                    group p by p.ToString() into g
                    where g.Count() > 2//出現1次以上的數字
                    select new
                    {
                        g.Key,
                        NumProducts = g.Count()
                    };
            foreach (var x in q)
            {
                Console.WriteLine(x.Key);//陣列中 每個數字出現的數量
            }
            Console.ReadLine();

        }








        // private int iTask = 1;
        void initTask()
        {
            iTask = 1;

        }
        //private async Task<bool> SpertroImageProcess(int i)
        private async Task<bool> SpertroImageProcess(byte[] FULLimgbuffer, int start_pixel)
        {
            await Task.Run(() =>
            {

                FormUpdata formcontrl = new FormUpdata(FormUpdataMethod);
                

                switch (iTask)
                {
                    case 1:
                        //   MessageBox.Show("開始第一階段?");
                        Dark_Intensity.Clear();
                        Original_Intensity.Clear();
                        Pure_Intensity.Clear();
                        SG_Intensity.Clear();
                        Save_Laser_Intensity.Clear();

                        WaveLength = new List<double>(8);
                        Poly_Coefs = new List<double>(5);
                        Pixel_Max = new List<double>(8);
                        Intensity_Max = new List<double>(8);
                        SD = new List<double>(8);
                        pFWHM = new List<double>(8);
                        wFWHM = new List<double>(8);
                        Laser = 1;
                        avg = 0;



                        //將所有相機參數調為預設值
                        SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 0;CHAN 2;LAS:OUT 0;CHAN 3;LAS:OUT 0;CHAN 4;LAS:OUT 0;CHAN 5;LAS:OUT 0;CHAN 6;LAS:OUT 0;CHAN 7;LAS:OUT 0;CHAN 8;LAS:OUT 0;\r\n"));
                        Thread.Sleep(1500);
                        SendData(Encoding.ASCII.GetBytes("BEEP 2;CHAN " + numericUpDown1.Value + ";LAS:OUT 1\r\n"));
                        Thread.Sleep(3000);
                        iTask = 100;
                        break;
                    case 100://Auto-scaling //判斷在初始dg時是否過曝
                        if (checkBox4.Checked) { iTask = 200; break; } //跳過
                        /*  int mm = FULLimgbuffer.Max();
                          if (mm<255) //沒過曝，繼續走
                          {
                              //  MessageBox.Show("初始Autoscaling已完成!");
                              iTask = 150;
                          }
                          else
                          {
                              Set_Dg_half();

                              iTask = 100;//100
                          }*/
                        iTask = 150;
                        Thread.Sleep(100);


                        if (Calibrate_DG(FULLimgbuffer))
                        {
                            //  MessageBox.Show("初始Autoscaling已完成!");
                            iTask = 200;
                        }
                        else
                        {
                            iTask = 100;//100
                        }

                        Thread.Sleep(100);


                        break;
                    case 150: //s
                              //   Dictionary<string , List<double>> result_buffer = new Dictionary<string, List<double>>;
                              //Step 0 找到最大值的位置 紀錄此時的"dg,max",index
                        /*   List<double> result_buffer = new List<double>();
                           List<double> dg_set = new List<double>();
                           List<double> Max_Intensity = new List<double>();
                           double Index_of_Max = 0;


                           result_buffer = Smart_Calibrate_DG(FULLimgbuffer, Index_of_Max); //輸出一個dg值
                           dg_set.Add(result_buffer[0]);
                           Max_Intensity.Add(result_buffer[1]);
                           Index_of_Max = result_buffer[2];

                           //Step 1 根據最大值的位置 找到dg/2後的dg,"max",index
                           result_buffer = Smart_Calibrate_DG(FULLimgbuffer, Index_of_Max); //輸出一個dg值
                           dg_set.Add(result_buffer[0]);
                           Max_Intensity.Add(result_buffer[1]);


                           List<double> Auto_Scaling_Coef = Math_Methods.Polynomial_Fitting(Max_Intensity, dg_set, 1);

                          int show_new_dg =  Convert.ToInt32(Auto_Scaling_Coef[0] + Auto_Scaling_Coef[1] * 190);
                           Set_Dg(Convert.ToInt32(Auto_Scaling_Coef[0]+ Auto_Scaling_Coef[1]*190));*/

                        iTask = 200;
                        break;

                    case 200:  //RoiScan

                        iTask = 300;
                        ROI_Y = Step_1.RoiScan(BufferToImage(FULLimgbuffer))["y"];


                        break;
                    case 300:


                        iTask = 400;


                        break;


                    case 400://Step2起始
                             //  MessageBox.Show("開始第二階段?");
                        Context = "λ(nm)\tPixel\tΔP(FWHM)\tΔλ(FWHM)\r\n";
                        iTask = 500;
                        WaveLength.Add(Math.Round(double.Parse(w1.Text, CultureInfo.InvariantCulture.NumberFormat), 2));

                        WaveLength.Add(Math.Round(double.Parse(w2.Text, CultureInfo.InvariantCulture.NumberFormat), 2));


                        WaveLength.Add(Math.Round(double.Parse(w3.Text, CultureInfo.InvariantCulture.NumberFormat), 2));

                        WaveLength.Add(Math.Round(double.Parse(w4.Text, CultureInfo.InvariantCulture.NumberFormat), 2));
                        if (ud_LaserCount.Value == 4) break;
                        WaveLength.Add(Math.Round(double.Parse(w5.Text, CultureInfo.InvariantCulture.NumberFormat), 2));
                        if (ud_LaserCount.Value == 5) break;
                        WaveLength.Add(Math.Round(double.Parse(w6.Text, CultureInfo.InvariantCulture.NumberFormat), 2));
                        if (ud_LaserCount.Value == 6) break;
                        WaveLength.Add(Math.Round(double.Parse(w7.Text, CultureInfo.InvariantCulture.NumberFormat), 2));
                        if (ud_LaserCount.Value == 7) break;
                        WaveLength.Add(Math.Round(double.Parse(w8.Text, CultureInfo.InvariantCulture.NumberFormat), 2));
                        break;

                    //  Step = 2; //之後要調回0


                    case 500: //刷新 //全部關掉
                        SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 0;CHAN 2;LAS:OUT 0;CHAN 3;LAS:OUT 0;CHAN 4;LAS:OUT 0;CHAN 5;LAS:OUT 0;CHAN 6;LAS:OUT 0;CHAN 7;LAS:OUT 0;CHAN 8;LAS:OUT 0;\r\n"));
                        //for (int i = 0; i < 100000; i++) for (int j = 0; j < 10000; j++) ;

                        /*     Console.WriteLine("sync: Starting");
                             Console.WriteLine("sync: Running for {0} seconds", DateTime.Now.Subtract(start).TotalSeconds);*/
                        Thread.Sleep(1000);
                        //     Console.WriteLine("sync: End for {0} seconds", DateTime.Now.Subtract(start).TotalSeconds);
                        if (isGetDarkChart == false)
                        {
                            Directory.CreateDirectory(@"Screenshot\" + this_SC_ID);
                            chartScreenShot getDark = new chartScreenShot(chartScreenshot);
                            this.Invoke(getDark, new Object[] { "_Dark_" });
                            isGetDarkChart = true;
                        }
                        iTask = 550;


                        break;
                    case 550: //刷新
                        //for (int i = 0; i < 100000; i++) for (int j = 0; j < 10000; j++) ;
                        Thread.Sleep(1000);


                        iTask = 600;
                        if (AllinOneMode.Checked)//進入一次八根
                        {
                            iTask = 601;
                        }
                        if (checkBox5.Checked)
                        { iTask = 600; }
                        break;
                    case 600://開啟單根雷射

                        if (Laser != 1 && Laser != 4 && Laser <= 8)
                        {
                            SendData(Encoding.ASCII.GetBytes("BEEP 2;CHAN " + Laser + ";LAS:OUT 1\r\n"));
                            Thread.Sleep(3000);

                            iTask = 625;

                        }
                        else if (Laser == 4)
                        {
                            Laser++;
                            iTask = 600;
                        }
                        else if (Laser == 1) //暫時排除L1
                        {
                            Laser++;
                            iTask = 600;
                        }


                        break;
                    case 601://開啟每根雷射，先確認該根雷射是否被選擇
                        if (AllinOneMode.Checked)
                        {
                            if (isch1)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch2)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 2;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch3)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 3;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch4)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 4;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch5)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 5;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch6)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 6;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch7)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 7;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch8)
                            { SendData(Encoding.ASCII.GetBytes("CHAN 8;LAS:OUT 1;\r\n")); Thread.Sleep(150); }
                            if (isch1 == false && isch2 == false && isch3 == false && isch4 == false && isch5 == false && isch6 == false && isch7 == false && isch8 == false)
                            {
                                SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 1;CHAN 2;LAS:OUT 1;CHAN 3;LAS:OUT 1;CHAN 5;LAS:OUT 1;CHAN 6;LAS:OUT 1;CHAN 7;LAS:OUT 1;CHAN 8;LAS:OUT 1;\r\n")); Thread.Sleep(150);
                            }

                        }
                        else
                        {
                            SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 1;CHAN 2;LAS:OUT 1;CHAN 3;LAS:OUT 1;CHAN 4;LAS:OUT 1;CHAN 5;LAS:OUT 1;CHAN 6;LAS:OUT 1;CHAN 7;LAS:OUT 1;CHAN 8;LAS:OUT 1;\r\n"));
                            Thread.Sleep(3000);
                        }
                        iTask = 625;

                        break;

                      
                    case 625://101->auto-scaling     //汞氬燈 與白燈 進入點
                        
                        if (isHg_Ar)
                        {
                            Set_Dg(220);
                            Set_Back(3);
                            if (isFirtCacular == 1)
                            {
                                

                                Dark_Intensity = RealTime_Original_Intensity;
                            }
                            ROI_Y = Step_1.RoiScan(BufferToImage(FULLimgbuffer))["y"];
                            Thread.Sleep(1000);
                            if (isGet_Ar == false)
                            {
                                Set_Gamma(300);                              
                            }
                            else
                            {
                                if (gamma_number > 300)
                                {
                                    MessageBox.Show("gamma已達最高");
                                }
                                else
                                {
                                    
                                    Set_Gamma(gamma_number);
                                   
                                }
                            }
                            iTask = 630;
                        }
                        else if (onlyWhite)
                        {
                            Set_Dg(255);
                            Set_Gamma(100);
                            if (back_number > 3)
                            {
                                MessageBox.Show("請更換減光片");
                            }
                            else
                            {
                                Set_Back(back_number);
                            }
                            if (isGetDarkChart == false)
                            {
                                Directory.CreateDirectory(@"Screenshot\" + this_SC_ID);
                                chartScreenShot getDark = new chartScreenShot(chartScreenshot);
                                this.Invoke(getDark, new Object[] { "_White_Dark_" });
                                isGetDarkChart = true;
                            }
                            if (isGetROI == false)
                            {
                                Dark_Intensity = RealTime_Original_Intensity;
                                ROI_Y = Step_1.RoiScan(BufferToImage(FULLimgbuffer))["y"];
                                isGetROI = true;
                            }
                            iTask = 630;
                        }
                        else
                        {
                          
                             Set_Dg(220);
                             Set_Back(3);
                           
                            if (gamma_number > 300)
                            {
                                MessageBox.Show("gamma已達最高");
                            }
                            else
                            {
                                Set_Gamma(gamma_number);
                            }
                                                 
                            iTask = 630;
                        }
                        Original_Intensity = RealTime_Original_Intensity;
                        Thread.Sleep(1500);
                        break;

                    case 630:
                        if (isGet_Ar == false)
                        {
                            Ar_Intensity = RealTime_Original_Intensity;
                            Set_Gamma(100);
                            isGet_Ar = true;
                        }
                        dg_set.Clear();
                        Max_Intensity.Clear();
                        //Step 0 找到最大值的位置 紀錄此時的"dg,max",index
                        result_buffer = Smart_Calibrate_DG_Intensity(RealTime_Original_Intensity, Index_of_Max); //輸出一個dg值
                        if (result_buffer[1] >= 250)
                        {
                            iTask = 630;
                            Index_of_Max = 0;
                        }
                        else
                        {
                            dg_set.Add(result_buffer[0]);
                            Max_Intensity.Add(result_buffer[1]);
                            Index_of_Max = result_buffer[2];
                            //  Set_Dg_half();//-------------
                            iTask = 635;
                        }
                        Thread.Sleep(1500);
                        break;

                    case 635:
                       
                        if (onlyWhite)
                        {
                            double max_ = Convert.ToDouble(textBox8.Text);
                            int Goal_intensity = Convert.ToInt32(256 * (max_ / 100));
                            //Step 1 根據最大值的位置 找到dg/2後的dg,"max",index
                            result_buffer = Smart_Calibrate_DG_Intensity(RealTime_Original_Intensity, Index_of_Max); //輸出一個dg值
                            dg_set.Add(result_buffer[0]);
                            Max_Intensity.Add(result_buffer[1]);

                            List<double> Auto_Scaling_Coef = Math_Methods.Polynomial_Fitting(Max_Intensity, dg_set, 1);

                            int show_new_dg = Convert.ToInt32(Auto_Scaling_Coef[0] + Auto_Scaling_Coef[1] * Goal_intensity);
                            if (show_new_dg >= 254)
                            {
                                back_number++;
                                iTask = 625;
                            }
                            else
                            {
                                Set_Dg(show_new_dg);
                                iTask = 638;
                            }
                        }
                        else
                        {
                            double max_ = Convert.ToDouble(textBox8.Text);
                            int Goal_intensity = Convert.ToInt32(256 * (max_ / 100));
                            //Step 1 根據最大值的位置 找到dg/2後的dg,"max",index
                            result_buffer = Smart_Calibrate_DG_Intensity(RealTime_Original_Intensity, Index_of_Max); //輸出一個dg值
                            dg_set.Add(result_buffer[0]);
                            Max_Intensity.Add(result_buffer[1]);
                            List<double> Auto_Scaling_Coef = Math_Methods.Polynomial_Fitting(Max_Intensity, dg_set, 1);
                            int show_new_dg = Convert.ToInt32(Auto_Scaling_Coef[0] + Auto_Scaling_Coef[1] * Goal_intensity);
                            if (show_new_dg >= 254)
                            {
                                gamma_number += 50;
                                iTask = 625;
                            }
                            else
                            {
                                Set_Dg(show_new_dg);
                                isGet_Ar = false;
                                iTask = 640;
                            }
                        }
                        Index_of_Max = 0;
                        isAutoScaling_FisrtStep_Complete = false;
                        Thread.Sleep(1500);
                        break;

                    case 638:
                        iTask = 640;
                        //-------------------------------白光Noise計算----------------------------------
                        List<double> White_Noise = new List<double>();
                        double mean_of_Noise = 0;
                        double sqr_Noise = 0;
                        double White_RMS_of_DarkNoise = 0;

                        for (int i = 0; i < 50; i++)
                        {
                            White_Noise.Add(RealTime_Original_Intensity[i]);
                        }
                        for (int i = 0; i < White_Noise.Count(); i++)
                        {
                            mean_of_Noise += White_Noise[i];
                        }
                        mean_of_Noise = mean_of_Noise / White_Noise.Count();

                        for (int i = 0; i < White_Noise.Count(); i++)
                        {
                            sqr_Noise += Math.Pow(White_Noise[i], 2);
                        }
                        White_RMS_of_DarkNoise = Math.Sqrt(sqr_Noise / White_Noise.Count());
                        White_Imax_Divided_By_Dark = RealTime_Original_Intensity.Max()/ mean_of_Noise;
                        White_Imax_Divided_By_DarkNoise = RealTime_Original_Intensity.Max() / White_RMS_of_DarkNoise;

                        break;


                    case 640: //將CHART 光譜圖存成圖檔
                        iTask = 650;
                        Directory.CreateDirectory(@"Screenshot\" + this_SC_ID);
                        chartScreenShot css = new chartScreenShot(chartScreenshot);


                        if (isHg_Ar)
                        {

                            //chart_original.SaveImage(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_Hg_Ar_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg", ImageFormat.Bmp);
                            this.Invoke(css, new Object[] { "_Hg_Ar_" });
                        }
                        else if (onlyWhite)
                        {                          
                            this.Invoke(css, new Object[] { "_White_" });
                        }
                        else if (checkBox5.Checked && isHg_Ar == false)
                        {
                            this.Invoke(css, new Object[] { $"_Laser{Laser}_" });
                        }
                        else
                        {

                            //chart_original.SaveImage(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_Laser_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg", ImageFormat.Bmp);
                            this.Invoke(css, new Object[] { "Laser" });
                        }

                        break;



                    case 650: //取Original N次平均

                        if (avg < numericUpDown2.Value)
                        {

                            Original_Intensity = Math_Methods.List_Add(Original_Intensity, RealTime_Original_Intensity);

                            Thread.Sleep(100);
                            avg++;
                            iTask = 650;
                        }
                        else
                        {
                            Original_Intensity = Math_Methods.List_Div(Original_Intensity, avg);
                            isAutoScaling_END = true;
                            isAvgEnd = true;
                            avg = 0;
                            iTask = 701;
                        }
                       //再取一次新參數後的光譜平均

                        break;

                    case 700:  //關單根雷射  SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 0;CHAN 2;LAS:OUT 0;CHAN 3;LAS:OUT 0;CHAN 4;LAS:OUT 0;CHAN 5;LAS:OUT 0;CHAN 6;LAS:OUT 0;CHAN 7;LAS:OUT 0;CHAN 8;LAS:OUT 0;\r\n"));
                        //for (int i = 0; i < 100000; i++) for (int j = 0; j < 10000; j++) ;
                        SendData(Encoding.ASCII.GetBytes("LAS:OUT 0\r\n"));

                        Thread.Sleep(1000);

                        iTask = 750;
                        break;

                    case 701:  //關每根雷射  
                        if (isHg_Ar)
                        { iTask = 750; }
                        else if (onlyWhite)
                        { iTask = 790; }
                        else
                        {
                            SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 0;CHAN 2;LAS:OUT 0;CHAN 3;LAS:OUT 0;CHAN 4;LAS:OUT 0;CHAN 5;LAS:OUT 0;CHAN 6;LAS:OUT 0;CHAN 7;LAS:OUT 0;CHAN 8;LAS:OUT 0;\r\n"));

                            Thread.Sleep(2000);
                            iTask = 750;
                        }
                        break;

                    case 750:  //取暗光譜+減暗光譜
                        List<string> original_int = new List<string>();

                        if (AllinOneMode.Checked && isHg_Ar == false)
                        {

                            Dark_Intensity = RealTime_Original_Intensity;
                            Pure_Intensity = Math_Methods.Remove_BaseLine(Original_Intensity, Dark_Intensity);
                            
                            iTask = 775;
                        }
                        else if (isHg_Ar)
                        {
                            Ar_Intensity = Math_Methods.Remove_BaseLine(Ar_Intensity, Dark_Intensity);
                            Pure_Intensity = Math_Methods.Remove_BaseLine(Original_Intensity, Dark_Intensity);
                            iTask = 775;
                        }
                       
                        else if (checkBox5.Checked && isHg_Ar == false && onlyWhite == false)
                        {
                            //-------------------------------------------存RAW_DATA-------------------------------------------------------------
                            Dark_Intensity = RealTime_Original_Intensity;
                            List<string> original = new List<string>();
                            for (int j = 0; j < Original_Intensity.Count; j++)
                            {
                                original.Add(Original_Intensity[j].ToString());
                            }
                            List<string> dark_Int = new List<string>();
                            for (int t = 0; t < Dark_Intensity.Count; t++)
                            {
                                dark_Int.Add(Dark_Intensity[t].ToString());
                            }
                            Directory.CreateDirectory(@"RawData\Auto_Save\" + this_SC_ID);
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + $"_Laser{Laser}_" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", original.ToArray());
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + $"_Dark_of_Laser_{Laser}" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", dark_Int.ToArray());
                            //---------------------------------------------------------------------------------------------------------------------
                            int Max_Index = 0;
                            bool isend1 = false;
                            bool isend2 = false;
                            bool end = false;
                            double avg = 0;
                            int i = 1;
                            int right_bounded = 0;
                            int left_bounded = 0;
                            Pure_Intensity = Math_Methods.Remove_BaseLine(Original_Intensity, Dark_Intensity);
                            Save_Laser_Intensity = Pure_Intensity;
                            avg = Dark_Intensity.Average();
                            if (Laser == 2) //暫時改2
                            {
                                Laser_Intensity = Pure_Intensity;
                            }
                            else
                            {
                                Max_Index = Save_Laser_Intensity.IndexOf(Save_Laser_Intensity.Max());
                                while (end == false)
                                {
                                    if (Save_Laser_Intensity[Max_Index + i] < avg && isend1 == false)
                                    {
                                        right_bounded = Max_Index + i;
                                        isend1 = true;
                                    }
                                    else if (Save_Laser_Intensity[Max_Index - i] < avg && isend2 == false)
                                    {
                                        left_bounded = Max_Index - i;
                                        isend2 = true;
                                    }
                                    else if (isend1 && isend2)
                                    {
                                        end = true;
                                    }
                                    else
                                    { i++; }
                                }

                                Laser_Intensity.RemoveRange(left_bounded, right_bounded - left_bounded);

                                for (int item = left_bounded; item < right_bounded; item++)
                                {
                                    Laser_Intensity.Insert(item, Save_Laser_Intensity[item]);
                                }
                            }
                            if (Laser == 8)
                            {
                                Pure_Intensity = Laser_Intensity;
                                iTask = 775;
                            }
                            else
                            {
                                Laser++;
                                gamma_number = 100;
                                iTask = 600;
                            }
                        }
                        else
                        {
                            Dark_Intensity = RealTime_Original_Intensity;
                            Pure_Intensity = Math_Methods.Remove_BaseLine(Original_Intensity, Dark_Intensity);
                            if (checkBox3.Checked) Pure_Intensity = Math_Methods.SG_Fitting(Pure_Intensity, Convert.ToInt32(numericUpDown3.Value), Convert.ToInt32(textBox1.Text));
                            Gaus_Pixel_Intensity_Parameter_Set = Math_Methods.Auto_Gaussian(Pure_Intensity);
                            iTask = 775;
                        }
                        break;

                    case 775:  //計算Noise RMS
                        int c = 0;
                        //double mean_of_Noise = 0;
                        double sqr_of_Noise = 0;
                        for (int i = 100; i < 200; i++)
                        {
                            Noise[c] = Pure_Intensity[i];
                            c++;
                        }
                        /*for (int i = 0; i < Noise.Count(); i++)
                        {
                            mean_of_Noise += Noise[i];
                        }
                        mean_of_Noise = mean_of_Noise / Noise.Count();
                        */
                        for (int i = 0; i < Noise.Count(); i++)
                        {
                            sqr_of_Noise += Math.Pow(Noise[i], 2);
                        }
                        RMS_of_Noise = Math.Sqrt(sqr_of_Noise / Noise.Count());
                        Dynamic_Range = 255 / RMS_of_Noise;
                        max_intensity = Pure_Intensity.Max();
                        Pixel_of_maxIntensity = Pure_Intensity.IndexOf(max_intensity);

                        iTask = 780;


                        break;
                    case 780:  //篩選有用到的波長，並丟入FINDPEAKS中運算

                        iTask = 790;
                        deltaL.Clear();
                        Wavelength.Clear();
                        if (isHg_Ar)
                        {
                            #region 汞氬燈波長篩選
                            if (HgW1_CK.Checked)
                            { Wavelength.Add(Convert.ToDouble(HG_w1.Text)); } deltaL.Add(Convert.ToDouble(HG1.Text));
                            if (HgW2_CK.Checked)
                            { Wavelength.Add(Convert.ToDouble(HG_w2.Text)); } deltaL.Add(Convert.ToDouble(HG2.Text));
                            if (HgW3_CK.Checked)
                            { Wavelength.Add(Convert.ToDouble(HG_w3.Text)); } deltaL.Add(Convert.ToDouble(HG3.Text));
                            if (HgW4_CK.Checked)
                            { Wavelength.Add(Convert.ToDouble(HG_w4.Text)); } deltaL.Add(Convert.ToDouble(HG4.Text));
                            if (HgW5_CK.Checked)
                            { Wavelength.Add(Convert.ToDouble(HG_w5.Text)); } deltaL.Add(Convert.ToDouble(HG5.Text));
                            if (HgW6_CK.Checked)
                            { Wavelength.Add(Convert.ToDouble(HG_w6.Text)); } deltaL.Add(Convert.ToDouble(HG6.Text));
                            #endregion
                        }
                        else {
                            #region 雷射波長篩選
                            Wavelength.Clear();
                            if (isch1)
                            { Wavelength.Add(Convert.ToDouble(w1.Text)); deltaL.Add(Convert.ToDouble(L1.Text)); }
                            if (isch2)
                            { Wavelength.Add(Convert.ToDouble(w2.Text)); deltaL.Add(Convert.ToDouble(L2.Text)); }
                            if (isch3)
                            { Wavelength.Add(Convert.ToDouble(w3.Text)); deltaL.Add(Convert.ToDouble(L3.Text)); }
                            if (isch4)
                            { Wavelength.Add(Convert.ToDouble(w4.Text)); deltaL.Add(Convert.ToDouble(L4.Text)); }
                            if (isch5)
                            { Wavelength.Add(Convert.ToDouble(w5.Text)); deltaL.Add(Convert.ToDouble(L5.Text)); }
                            if (isch6)
                            { Wavelength.Add(Convert.ToDouble(w6.Text)); deltaL.Add(Convert.ToDouble(L6.Text)); }
                            if (isch7)
                            { Wavelength.Add(Convert.ToDouble(w7.Text)); deltaL.Add(Convert.ToDouble(L7.Text)); }
                            if (isch8)
                            { Wavelength.Add(Convert.ToDouble(w8.Text)); deltaL.Add(Convert.ToDouble(L8.Text)); }
                            if (isch1 == false && isch2 == false && isch3 == false && isch4 == false && isch5 == false && isch6 == false && isch7 == false && isch8 == false)
                            {
                                Wavelength.Add(Convert.ToDouble(w1.Text)); Wavelength.Add(Convert.ToDouble(w2.Text));
                                Wavelength.Add(Convert.ToDouble(w3.Text)); Wavelength.Add(Convert.ToDouble(w5.Text));
                                Wavelength.Add(Convert.ToDouble(w6.Text)); Wavelength.Add(Convert.ToDouble(w7.Text));
                                Wavelength.Add(Convert.ToDouble(w8.Text));
                                deltaL.Add(Convert.ToDouble(L1.Text)); deltaL.Add(Convert.ToDouble(L2.Text));
                                deltaL.Add(Convert.ToDouble(L3.Text)); deltaL.Add(Convert.ToDouble(L5.Text));
                                deltaL.Add(Convert.ToDouble(L6.Text)); deltaL.Add(Convert.ToDouble(L7.Text));
                                deltaL.Add(Convert.ToDouble(L8.Text));
                            }
                            #endregion
                        }
                        waveCount = Wavelength.Count;
                        if (isHg_Ar)
                        {
                            Dictionary<string, List<double>> Gaus_FWHL_Coef_Set = new Dictionary<string, List<double>>();
                            Gaus_FWHL_Coef_Set = Math_Methods.Hg_FindPeak_And_Gaussian(Pure_Intensity, Wavelength, Convert.ToDouble(textBox12.Text), Convert.ToDouble(textBox13.Text), Convert.ToDouble(textBox14.Text));
                            Gaus_FWHL_Coef_Set_1 = Math_Methods.Ar_FindPeak_And_Gaussian(Gaus_FWHL_Coef_Set, Pure_Intensity, Ar_Intensity, Wavelength, Convert.ToDouble(textBox12.Text), Convert.ToDouble(textBox13.Text), Convert.ToDouble(textBox14.Text));
                        }
                        else
                        {
                            Gaus_FWHL_Coef_Set_1 = Math_Methods.FindPeak_And_Gaussian(Pure_Intensity, Wavelength, Convert.ToDouble(textBox12.Text), Convert.ToDouble(textBox13.Text), Convert.ToDouble(textBox14.Text));
                        }


                        break;

                    case 790://自動將光譜存成TXT
                        iTask = 800;
                        List<string> original_ = new List<string>();
                        List<string> dark_ = new List<string>();
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            original_.Add(RealTime_Original_Intensity[i].ToString());
                            dark_.Add(Dark_Intensity[i].ToString());
                          
                        }
                        Directory.CreateDirectory(@"RawData\Auto_Save\" + this_SC_ID);

                        if (isHg_Ar)
                        {
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "_Hg_Ar_" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", original_.ToArray());
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "_Dark_of_Hg_Ar_" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", dark_.ToArray());
                        }
                        else if (onlyWhite)
                        {
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "_While_" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", original_.ToArray());
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "_Dark_of_White_" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", dark_.ToArray());
                            iTask = 1200;
                        }
                        else
                        {
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "_Laser_" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", original_.ToArray());
                            File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "_Dark_of_Laser" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", dark_.ToArray());
                        }
                        break;

                   


                    case 800: //將計算出來的結果參數 丟進 PUBLIC 變數中
                        if (AllinOneMode.Checked)
                        {
                            pFWHM = Gaus_FWHL_Coef_Set_1["FWHL"];//pFWHM 
                            Poly_Coefs = Gaus_FWHL_Coef_Set_1["coef"];
                            pixel_M = Gaus_FWHL_Coef_Set_1["pixel_M"];
                        }
                        else if (checkBox5.Checked)
                        {
                            pFWHM = Gaus_FWHL_Coef_Set_1["FWHL"];//pFWHM 
                            Poly_Coefs = Gaus_FWHL_Coef_Set_1["coef"];
                            pixel_M = Gaus_FWHL_Coef_Set_1["pixel_M"];
                        }
                        else
                        {
                            Pixel_Max.Add(Gaus_Pixel_Intensity_Parameter_Set["Parameters"][0]);//pFWHM
                            Intensity_Max.Add(Gaus_Pixel_Intensity_Parameter_Set["Parameters"][1]);//pFWHM 
                            SD.Add(Gaus_Pixel_Intensity_Parameter_Set["Parameters"][2]);//pFWHM 
                            pFWHM.Add(Gaus_Pixel_Intensity_Parameter_Set["Parameters"][3]);//pFWHM 
                            Context += Math.Round(WaveLength[Laser - 1], 2) + "\t" + Math.Round(Pixel_Max[Laser - 1], 2) + "\t" + Math.Round(pFWHM[Laser - 1], 2) + "\t                 " + "?" + Laser + "\r\n";
                        }


                        iTask = 900;
                        break;



                    case 900: //多項式擬和

                        Laser++;
                        if (AllinOneMode.Checked) { //全波段時
                            iTask = 1000;

                        }
                        else if(checkBox5.Checked)
                        { 
                            iTask = 1000;
                        }
                        else if (Laser >= ud_LaserCount.Value + 1)
                        {
                            iTask = 1000;
                            //      Step = 0; //之後調成3?
                        }
           

                        break;

                    case 1000: //Step3 Poly
                        if (AllinOneMode.Checked)
                        {
                            iTask = 1001;
                        }
                        else if (checkBox5.Checked)
                        {
                            iTask = 1001;
                        }
                        else
                        {
                            Poly_Coefs.Clear();
                            wFWHM.Clear();

                            Poly_Coefs = Math_Methods.Polynomial_Fitting(Pixel_Max, WaveLength, Convert.ToInt32(ud_Poly.Value));

                            wFWHM = Math_Methods.getLamdaFWHM(pFWHM, Poly_Coefs, Pixel_Max);
                            //依序把?改成wFWHM
                            for (int L = 1; L <= ud_LaserCount.Value; L++) Context = Context.Replace("?" + L, Math.Round(wFWHM[L - 1], 5).ToString() + " nm");

                            //新增擬和結果
                            //      Context += "----------------------------------------------------------------\na4\ta3\ta2\ta1\ta0\n" + Math.Round(Poly_Coefs[0], 2) + "\t" + Math.Round(Poly_Coefs[1], 2) + "\t" + Math.Round(Poly_Coefs[2], 2) + "\t" + Math.Round(Poly_Coefs[3], 2) + "\t" + Math.Round(Poly_Coefs[4], 2)"\n" ;
                            Context += "----------------------------------------------------------------\na0 = " + Poly_Coefs[0] + "\na1 = " + Poly_Coefs[1] + "\na2 = " + Poly_Coefs[2] + "\na3 = " + Poly_Coefs[3] + "\na4 = " + Poly_Coefs[4] + "\n";
                            //   this.Invoke(formcontrl, ROI_Y / scaleY, 0, Context);
                            //   this.Invoke(formcontrl, DrawCanvas.Top, 0, Context, Result);
                            iTask = 1100;
                        }

                        break;

                    case 1001: //求 Δλm(ΔλFWHM)

                        wFWHM.Clear();
                        wFWHM = Math_Methods.getLamdaFWHM(pFWHM, Poly_Coefs, pixel_M);
                        if (isHg_Ar)
                        { iTask = 1011; }
                        else { iTask = 1010; }
                        break;

                    case 1010:  //Δλsc

                        LambdaSC.Clear();
                        int k = 0;
                        int n = 0;
                        while (k != 9)
                        {
                            try
                            {
                                switch (k)
                                {
                                    case 0:

                                        if (isch1 && CH1_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L1.Text), 2)));
                                        }
                                        else if (isch1 && CH1_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 1;
                                        break;

                                    case 1:

                                        if (isch2 && CH2_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L2.Text), 2)));
                                        }
                                        else if (isch2 && CH2_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 2;
                                        break;

                                    case 2:

                                        if (isch3 && CH3_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L3.Text), 2)));
                                        }
                                        else if (isch3 && CH3_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 3;
                                        break;

                                    case 3:

                                        if (isch4 && CH4_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L4.Text), 2)));
                                        }
                                        else if (isch4 && CH4_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 4;
                                        break;

                                    case 4:

                                        if (isch5 && CH5_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L5.Text), 2)));
                                        }
                                        else if (isch5 && CH5_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 5;
                                        break;

                                    case 5:

                                        if (isch6 && CH6_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L6.Text), 2)));
                                        }
                                        else if(isch6 && CH6_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 6;
                                        break;

                                    case 6:

                                        if (isch7 && CH7_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L7.Text), 2)));
                                        }
                                        else if (isch7 && CH7_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 7;
                                        break;

                                    case 7:

                                        if (isch8 && CH8_CHECK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[k - n], 2) - Math.Pow(Convert.ToDouble(L8.Text), 2)));
                                        }
                                        else if (isch8 && CH8_CHECK.Checked == false)
                                        { n = n; }
                                        else { n++; }
                                        k = 9;
                                        iTask = 1012;
                                        break;
                                }
                            }
                            catch 
                            {
                                MessageBox.Show("找到的峰值數與預設不符"); iTask = 1300;k = 9; break; 
                            }
                        }

                        break;


                    case 1011: //Δλsc ( for hg )
                        
                        LambdaSC.Clear();
                        Poly_Coefs_of_Hg_Ar.Clear();
                        Poly_Coefs_of_Hg_Ar = Poly_Coefs;
                        int p = 0;
                        int r = 0;
                        while (p != 9)
                        {
                            try
                            {
                                switch (p)
                                {
                                    case 0:

                                        if (HgW1_CK.Checked && Hg1WFHM)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[p - r], 2) - Math.Pow(Convert.ToDouble(HG1.Text), 2)));
                                            //deltaL.Add(Convert.ToDouble(HG1.Text));
                                        }
                                        else { r++; }
                                        p = 1;
                                        break;

                                    case 1:

                                        if (HgW2_CK.Checked && Hg2WFHM)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[p - r], 2) - Math.Pow(Convert.ToDouble(HG2.Text), 2)));
                                            //deltaL.Add(Convert.ToDouble(HG2.Text));
                                        }
                                      
                                        else { r++; }
                                        p = 2;
                                        break;

                                    case 2:

                                        if (HgW3_CK.Checked && Hg3WFHM)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[p - r], 2) - Math.Pow(Convert.ToDouble(HG3.Text), 2)));
                                            //deltaL.Add(Convert.ToDouble(HG3.Text));
                                        }
                                       
                                        else { r++; }
                                        p = 3;
                                        break;

                                    case 3:

                                        if (HgW4_CK.Checked && Hg4WFHM)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[p - r], 2) - Math.Pow(Convert.ToDouble(HG4.Text), 2)));
                                            //deltaL.Add(Convert.ToDouble(HG4.Text));
                                        }
                                       
                                        else { r++; }
                                        p = 4;
                                        break;

                                    case 4:

                                        if (HgW5_CK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[p - r], 2) - Math.Pow(Convert.ToDouble(HG5.Text), 2)));
                                            //deltaL.Add(Convert.ToDouble(HG5.Text));
                                        }
                                        
                                        else { r++; }
                                        p = 5;
                                        break;

                                    case 5:

                                        if (HgW6_CK.Checked)
                                        {
                                            LambdaSC.Add(Math.Sqrt(Math.Pow(wFWHM[p - r], 2) - Math.Pow(Convert.ToDouble(HG6.Text), 2)));
                                            //deltaL.Add(Convert.ToDouble(HG6.Text));
                                        }
                                       
                                        else { r++; }
                                        p = 9;
                                        iTask = 1012;
                                        break;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("找到的峰值數與預設不符"); iTask = 1300; p = 9; break;
                            }
                        }

                            break;

                    case 1012: //將沒勾FWHM的波長找到的峰值所在的PIXEL也去掉
                        try
                        {
                            if (isHg_Ar == false)
                            {
                                RemoveIndex.Clear();
                                pixel_M_save.Clear();
                                int d = 0;
                                if (isch1 && CH1_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w1.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch2 && CH2_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w2.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch3 && CH3_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w3.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch4 && CH4_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w4.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch5 && CH5_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w5.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch6 && CH6_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w6.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch7 && CH7_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w7.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (isch8 && CH8_CHECK.Checked == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(w8.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                            }
                            else 
                            {
                                RemoveIndex.Clear();
                                pixel_M_save.Clear();
                                int d = 0;
                                if (HgW1_CK.Checked && Hg1WFHM == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(HG_w1.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (HgW2_CK.Checked && Hg2WFHM == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(HG_w2.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (HgW3_CK.Checked && Hg3WFHM == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(HG_w3.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                                if (HgW4_CK.Checked && Hg4WFHM == false)
                                {
                                    int i = Wavelength.IndexOf(Convert.ToDouble(HG_w4.Text));
                                    pixel_M_save.Add(pixel_M[i - d]);
                                    pixel_M.RemoveAt(i - d);
                                    RemoveIndex.Add(i);
                                    d++;
                                }
                            }
                        }
                        catch
                        {
                            MessageBox.Show("找到的峰值數與預設不符"); iTask = 1300; p = 9; break;
                        }
                        iTask = 1013;
                        break;

                    case 1013: //求ΔPsc
                        iTask = 1014;
                        try
                        {
                            DeltaPsc = Math_Methods.getDeltaPsc(LambdaSC, Poly_Coefs, pixel_M);
                        }
                        catch { MessageBox.Show("找到的峰值數與預設不符"); iTask = 1300; }

                        /*if (Laser_And_Hg_Ar_CHK.Checked && isHg_Ar)
                        {
                            for (int i = 0; i < Wavelength.Count; i++)
                            {
                                LambdaSC_save.Add(LambdaSC[i]);
                                DeltaPsc_save.Add(DeltaPsc[i]);
                                //Wavelength_save.Add(Wavelength[i]);
                            }
                        }*/
                        break;

                    case 1014:

                        Lamda_cal.Clear();
                        DIFF.Clear();
                        double x = 0;
                        double mean_of_SC = 0;
                        double y = 0;
                        double mean_of_SC_spot = 0;
                        double z = 0;
                        double s = 0;
                        double cal = 0;
                        //-----------------------------
                        double diff_rms = 0;

                       /* if (isHg_Ar)
                        {
                            //λcal
                            for (int i = 0; i < pixel_M.Count; i++)
                            {
                                cal = Poly_Coefs[3] * Math.Pow(pixel_M[i], 3) + Poly_Coefs[2] * Math.Pow(pixel_M[i], 2) + Poly_Coefs[1] * pixel_M[i] + Poly_Coefs[0];
                                Lamda_cal.Add(cal);
                            }
                            //DIFF
                            for (int i = 0; i < Wavelength.Count; i++)
                            {
                                DIFF.Add(Lamda_cal[i] - Wavelength[i]);
                            }
                            //DIFF_rms 
                            for (int i = 0; i < DIFF.Count; i++)
                            {
                                diff_rms += Math.Pow(DIFF[i], 2);
                            }
                            DIFF_rms = Math.Sqrt(diff_rms / DIFF.Count);
                        }*/
                       

                            for (int i = 0; i < LambdaSC.Count; i++)
                            {
                                mean_of_SC += LambdaSC[i];
                                mean_of_SC_spot += DeltaPsc[i] * 3.75;
                            }
                            mean_of_SC = mean_of_SC / LambdaSC.Count;
                            mean_of_SC_spot = mean_of_SC_spot / DeltaPsc.Count;

                            for (int i = 0; i < LambdaSC.Count; i++)
                            {
                                x += Math.Pow((LambdaSC[i] - mean_of_SC), 2);
                                y += Math.Pow((DeltaPsc[i] * 3.75 - mean_of_SC_spot), 2);
                                z += Math.Pow(LambdaSC[i], 2);
                                s += Math.Pow(DeltaPsc[i] * 3.75, 2);
                            }
                            x = x / LambdaSC.Count;
                            y = y / DeltaPsc.Count;
                            z = z / LambdaSC.Count;
                            s = s / DeltaPsc.Count;
                            LambdaSC_RMS = Math.Sqrt(z); //ΔλRMS
                            LambdaSC_Spot_RMS = Math.Sqrt(s);//ΔXRMS
                            LambdaSC_STD = Math.Sqrt(x);//Δλstd
                            LambdaSC_Spot_STD = Math.Sqrt(y);//ΔXstd
                            LambdaSC_STD_100 = (LambdaSC_STD / mean_of_SC) * 100;//Δλstd%
                            LambdaSC_Spot_STD_100 = (LambdaSC_Spot_STD / mean_of_SC_spot) * 100; //Δλstd%
                                                                                                 //------------------------------------------------------------------------
                            for (int i = 0; i < RemoveIndex.Count; i++)
                            {
                                pixel_M.Insert(RemoveIndex[i], pixel_M_save[i]);
                                LambdaSC.Insert(RemoveIndex[i], 0);
                                DeltaPsc.Insert(RemoveIndex[i], 0);

                            }
                            //SNR-STRAYLIGHT-DYNAMIC_RANGE-----------------------------------
                            lamda_max_Intensity = Poly_Coefs[3] * Math.Pow(Pixel_of_maxIntensity, 3) + Poly_Coefs[2] * Math.Pow(Pixel_of_maxIntensity, 2) + Poly_Coefs[1] * Pixel_of_maxIntensity + Poly_Coefs[0];
                            SNR = max_intensity / RMS_of_Noise;
                            Stray_light = (1 / SNR) * 100;

                            //λcal
                            for (int i = 0; i < pixel_M.Count; i++)
                            {
                                cal = Poly_Coefs[3] * Math.Pow(pixel_M[i], 3) + Poly_Coefs[2] * Math.Pow(pixel_M[i], 2) + Poly_Coefs[1] * pixel_M[i] + Poly_Coefs[0];
                                Lamda_cal.Add(cal);
                            }
                            //DIFF
                            for (int i = 0; i < Wavelength.Count; i++)
                            {
                                DIFF.Add(Lamda_cal[i] - Wavelength[i]);
                            }
                            //DIFF_rms 
                            for (int i = 0; i < DIFF.Count; i++)
                            {
                                diff_rms += Math.Pow(DIFF[i], 2);
                            }
                            DIFF_rms = Math.Sqrt(diff_rms / DIFF.Count);
                        
                            iTask = 1020;                       
                        break;


                    case 1020:
                        if (Laser_And_Hg_Ar_CHK.Checked == false)
                        {
                            report = "λ , ΔλL(nm)" + "\t" + "λcal"+ "\t" +"Diff"+ "\t" +"Diff_rms"+ "\t" + "Pixel" + "\t" + "ΔP" + "\t" + "Δxm" + "\t" + "Δλm" + "\t" + "Δλsc" + "\t" + "ΔPsc" + "\t" + "ΔXsc" + "  " + "Δλrms" + "  " + "  " + "Δλstd" + "  " + "  " + "Δλstd%" + "  " + "  " + "ΔXrms" + "  " + "  " + "ΔXstd" + "  " + "  " + "ΔXstd%" ;
                            if (isHg_Ar)
                            {
                                report += "\r\n" + "-----------------------------------------------------------------------------------------HG-----------------------------------------------------------------------------------------------------------------------\n";
                                for (int i = 0; i < Wavelength.Count; i++)
                                {
                                    report += "\r\n" + Wavelength[i].ToString("f2") + "  ,  " + deltaL[i].ToString("f2") //λ , ΔλL(nm)
                                        + "\t" + Lamda_cal[i].ToString("f2") //λcal
                                        + "\t" + DIFF[i].ToString("f2") //Diff
                                        + "\t" + DIFF_rms.ToString("f2") //Diff_rms
                                        + "\t" + pixel_M[i].ToString("f2") //Pixel
                                        + "\t" + pFWHM[i].ToString("f2") //ΔP
                                        + "\t" + (pFWHM[i] * 3.75).ToString("f2") //Δx
                                        + "\t" + wFWHM[i].ToString("f2") //Δλm
                                        + "\t" + LambdaSC[i].ToString("f2")//Δλsc
                                        + "\t" + DeltaPsc[i].ToString("f2") //ΔPsc
                                        + "\t" + (DeltaPsc[i] * 3.75).ToString("f2")//ΔXsc
                                        + "\t" + LambdaSC_RMS.ToString("f2")//ΔλRMS
                                        + "\t" + "  " + LambdaSC_STD.ToString("f2")//Δλstd
                                        + "\t" + "  " + LambdaSC_STD_100.ToString("f2")//ΔλRMS%
                                        + "\t" + "  " + "  " + LambdaSC_Spot_RMS.ToString("f2")//ΔXRMS
                                        + "\t" + "  " + "  " + LambdaSC_Spot_STD.ToString("f2")//ΔXstd
                                        + "\t" + "  " + "  " + LambdaSC_Spot_STD_100.ToString("f2");//ΔXRMS%

                                }
                            }
                            else
                            {
                                report += "\r\n" + "-----------------------------------------------------------------------------------------Laser--------------------------------------------------------------------------------------------------------------------\n";
                                for (int i = 0; i < Wavelength.Count; i++)
                                {
                                    report += "\r\n" + Wavelength[i].ToString("f2") + "  ,  " + deltaL[i].ToString("f2") //λ , ΔλL(nm)
                                        + "\t" + Lamda_cal[i].ToString("f2") //λcal
                                        + "\t" + DIFF[i].ToString("f2") //Diff
                                        + "\t" + DIFF_rms.ToString("f2") //Diff_rms
                                        + "\t" + pixel_M[i].ToString("f2") //Pixel
                                        + "\t" + pFWHM[i].ToString("f2") //ΔP
                                        + "\t" + (pFWHM[i] * 3.75).ToString("f2") //Δx
                                        + "\t" + wFWHM[i].ToString("f2") //Δλm
                                        + "\t" + LambdaSC[i].ToString("f2")//Δλsc
                                        + "\t" + DeltaPsc[i].ToString("f2") //ΔPsc
                                        + "\t" + (DeltaPsc[i] * 3.75).ToString("f2")//ΔXsc
                                        + "\t" + LambdaSC_RMS.ToString("f2")//ΔλRMS
                                        + "\t" + "  " + LambdaSC_STD.ToString("f2")//Δλstd
                                        + "\t" + "  " + LambdaSC_STD_100.ToString("f2")//ΔλRMS%
                                        + "\t" + "  " + "  " + LambdaSC_Spot_RMS.ToString("f2")//ΔXRMS
                                        + "\t" + "  " + "  " + LambdaSC_Spot_STD.ToString("f2")//ΔXstd
                                        + "\t" + "  " + "  " + LambdaSC_Spot_STD_100.ToString("f2");//ΔXRMS%

                                }
                            }
                            report += "\r\n" + "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"

                                            + "\r\n" + "a0 = " + Poly_Coefs[0]
                                            + "\r\n" + "a1 = " + Poly_Coefs[1]
                                            + "\r\n" + "a2 = " + Poly_Coefs[2]
                                            + "\r\n" + "a3 = " + Poly_Coefs[3]
                                            + "\r\n"
                                            + "\r\n";
                        }
                        else if (Laser_And_Hg_Ar_CHK.Checked)
                        {
                            report = "λ , ΔλL(nm)" + "\t" + "λcal" + "\t" + "Diff" + "\t" + "Diff_rms" + "\t" + "Pixel" + "\t" + "ΔP" + "\t" + "Δxm" + "\t" + "Δλm" + "\t" + "Δλsc" + "\t" + "ΔPsc" + "\t" + "ΔXsc" + "  " + "Δλrms" + "  " + "  " + "Δλstd" + "  " + "  " + "Δλstd%" + "  " + "  " + "ΔXrms" + "  " + "  " + "ΔXstd" + "  " + "  " + "ΔXstd%";
                            if (isHg_Ar)
                            {
                                report += "\r\n" + "-----------------------------------------------------------------------------------------HG-----------------------------------------------------------------------------------------------------------------------------------------------------------\n";
                                for (int i = 0; i < Wavelength.Count; i++)
                                {
                                    report += "\r\n" + Wavelength[i].ToString("f2") + "  ,  " + deltaL[i].ToString("f2") //λ , ΔλL(nm)
                                         + "\t" + Lamda_cal[i].ToString("f2") //λcal
                                         + "\t" + DIFF[i].ToString("f2") //Diff
                                         + "\t" + DIFF_rms.ToString("f2") //Diff_rms
                                         + "\t" + pixel_M[i].ToString("f2") //Pixel
                                         + "\t" + pFWHM[i].ToString("f2") //ΔP
                                         + "\t" + (pFWHM[i] * 3.75).ToString("f2") //Δx
                                         + "\t" + wFWHM[i].ToString("f2") //Δλm
                                         + "\t" + LambdaSC[i].ToString("f2")//Δλsc
                                         + "\t" + DeltaPsc[i].ToString("f2") //ΔPsc
                                         + "\t" + (DeltaPsc[i] * 3.75).ToString("f2")//ΔXsc
                                         + "\t" + LambdaSC_RMS.ToString("f2")//ΔλRMS
                                         + "\t" + "  " + LambdaSC_STD.ToString("f2")//Δλstd
                                         + "\t" + "  " + LambdaSC_STD_100.ToString("f2")//ΔλRMS%
                                         + "\t" + "  " + "  " + LambdaSC_Spot_RMS.ToString("f2")//ΔXRMS
                                         + "\t" + "  " + "  " + LambdaSC_Spot_STD.ToString("f2")//ΔXstd
                                         + "\t" + "  " + "  " + LambdaSC_Spot_STD_100.ToString("f2");//ΔXRMS%
                                }
                                report_Memory = report;
                            }
                            else
                            {
                                report += "\r\n" + "-----------------------------------------------------------------------------------------Laser------------------------------------------------------------------------------------------------------------------------------------\n";
                                for (int i = 0; i < Wavelength.Count; i++)
                                {
                                    report += "\r\n" + Wavelength[i].ToString("f2") + "  ,  " + deltaL[i].ToString("f2") //λ , ΔλL(nm)
                                         + "\t" + Lamda_cal[i].ToString("f2") //λcal
                                         + "\t" + DIFF[i].ToString("f2") //Diff
                                         + "\t" + DIFF_rms.ToString("f2") //Diff_rms
                                         + "\t" + pixel_M[i].ToString("f2") //Pixel
                                         + "\t" + pFWHM[i].ToString("f2") //ΔP
                                         + "\t" + (pFWHM[i] * 3.75).ToString("f2") //Δx
                                         + "\t" + wFWHM[i].ToString("f2") //Δλm
                                         + "\t" + LambdaSC[i].ToString("f2")//Δλsc
                                         + "\t" + DeltaPsc[i].ToString("f2") //ΔPsc
                                         + "\t" + (DeltaPsc[i] * 3.75).ToString("f2")//ΔXsc
                                         + "\t" + LambdaSC_RMS.ToString("f2")//ΔλRMS
                                         + "\t" + "  " + LambdaSC_STD.ToString("f2")//Δλstd
                                         + "\t" + "  " + LambdaSC_STD_100.ToString("f2")//ΔλRMS%
                                         + "\t" + "  " + "  " + LambdaSC_Spot_RMS.ToString("f2")//ΔXRMS
                                         + "\t" + "  " + "  " + LambdaSC_Spot_STD.ToString("f2")//ΔXstd
                                         + "\t" + "  " + "  " + LambdaSC_Spot_STD_100.ToString("f2");//ΔXRMS%

                                }
                                report += "\r\n" + "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"
                                            +"\r\n" + "a0 = " + Poly_Coefs[0]
                                            + "\r\n" + "a1 = " + Poly_Coefs[1]
                                            + "\r\n" + "a2 = " + Poly_Coefs[2]
                                            + "\r\n" + "a3 = " + Poly_Coefs[3]
                                            + "\r\n"
                                            + "\r\n";
                                report_Memory += "\r\n" + report;
                            }                           
                        }
                        iTask = 1030;
                        break;

                    case 1030: //Show Report
                        isend = true;
                        iTask = 1200;
                        isHg_Ar = false;
                        isLaser = false;
                        DialogResult isSHowReport = MessageBox.Show("是否出報表?" + "\r\n" + "若需等待其他結果請按" + @"""NO""" + "\r\n" + "若僅需一種結果請按" + @"""YES""", "是否出報表", MessageBoxButtons.YesNo);
                        if (isSHowReport.ToString() == "Yes")
                        {
                            Report report = new Report(this);
                            report.ShowDialog();
                            report_Memory = "";
                        }
                        
                        break;


                    case 1100: //Step4 Write to ISB

                        

                        iTask = 1200;

                        break;

                    case 1200://
                        if (onlyWhite) 
                        {
                            final_dg = textBox9.Text;
                            final_gamma = textBox10.Text;
                            final_back = textBox11.Text;
                            onlyWhite = false;
                            MessageBox.Show("測量結束");
                            
                        }
              
                        
                 
                        Pause = true;

                        iTask = 1300;
                        break;

                }
                this.Invoke(formcontrl, Convert.ToInt32(ROI_Y / scaleY), 0, Context, Result, Readble);

            });

            return true;
        }


        //Step2.各波長雷射峰值找尋
        private int iTask = 0;
        private int waveCount = 8;
        private int Laser = 1;
        private int avg = 0;
        public List<double> WaveLength = new List<double>(8);
        public List<double> Poly_Coefs = new List<double>(5);
        public List<double> Pixel_Max = new List<double>(8);
        public List<double> Intensity_Max = new List<double>(8);
        public List<double> SD = new List<double>(8);
        public List<double> pFWHM = new List<double>(8);
        public List<double> wFWHM = new List<double>(8);
        private bool isPeakLookingFinish = false;
        Dictionary<string, List<double>> Gaus_Pixel_Intensity_Parameter_Set = new Dictionary<string, List<double>>();
        Dictionary<string, List<double>> Gaus_FWHL_Coef_Set_1 = new Dictionary<string, List<double>>();
        private async Task<bool> Peak_Looking_Task(byte[] ROIimgbuffer, List<double> RealTime_Original_Intensity) //完成後回傳TRUE
        {

            Bitmap myBitmap = camera.GetBitmap();
            //Task.Run(() => ShowThreadInfo("Task"));
            await Task.Run(() =>
            {
                //影像處理程序 把Do While 刪除 寫在這
                //測試用
                //     Dictionary<string, List<double>> Gaus_Pixel_Intensity_Parameter_Set = new Dictionary<string, List<double>>();
                //      Step = 3;
                FormUpdata formcontrl = new FormUpdata(FormUpdataMethod);
                DateTime start = DateTime.Now;
                switch (iTask)
                {


                }
            });

            return true;
        }




        public Form1(int Gamma, int Back)
        {
            InitializeComponent();
            #region auto_change_form_size
            //代码增加ResizeEnd事件   //也可通过设计器增加ResizeEnd事件然后删掉此处代码

            this.ResizeEnd += new System.EventHandler(this.Form1_ResizeEnd);



            //设置最小大小为窗体打开的初始大小

            //this.MinimumSize = this.Size;



            //取得Form中控件中初始的字体大小 //若此From中字体不统一，则此处为数组。

            FontSizeMin = this.label2.Font.Size; //2020-03-23 19:21:50 更新



            //获得窗体改变前的Form的宽度与高度 //赋初始值

            iForm_ResizeBefore_Width = this.Size.Width;

            iForm_ResizeBefore_Height = this.Size.Height;
            #endregion
            /*
            DrawCanvas.Parent = CCDImage;
            ROIImage.Parent = panel1;
            //初始化 ROI位置
            ROI_Point.X = 0;
            ROI_Point.Y = 0;
            DrawCanvas.Left = ROI_Point.X;
            DrawCanvas.Top = ROI_Point.Y;
            
            CCDImage.Width = 640;
            CCDImage.Height = 480;
            panel1.Width = 640;
            panel1.Height = 480;
            ROIImage.Width = 640;
            ROIImage.Height = 480;
            
            CCDImage.Width = 480;
            CCDImage.Height = 360;
            panel1.Width = 480;
            panel1.Height = 360;
            ROIImage.Width = 480;
            ROIImage.Height = 360;

            DrawCanvas.Width = CCDImage.Width;
            DrawCanvas.Height = 10;
            */

            //相機參數調整
            //設定
            set_camera_prop("exp", -2);
            set_camera_prop("bright", 500);
            set_camera_prop("con", 0);

            set_camera_prop("hue", -2000);
            set_camera_prop("sat", 0);
            set_camera_prop("sharp", 1);
            //gamma
            set_camera_prop("white", 2800);
            //back
            //gain
            set_camera_prop("gain", 200);

            Camera_Parameters.Clear();

            Camera_Parameters.Add("gamma", Gamma);
            Camera_Parameters.Add("back", Back);

            set_camera_prop("gamma", Camera_Parameters["gamma"]);
            set_camera_prop("back", Camera_Parameters["back"]);


        }

        private void btnStart_Click(object sender, EventArgs e)
        {


            if (bCameraLive == false)
            {
                bCameraLive = true;
                camera.Start();

                //var bmp = camera.GetBitmap();
                btnStart.Text = "Stop";
                button9.Enabled = false;
                // show image in PictureBox.
                timer1.Start();

            }
            else
            {
                bCameraLive = false;
                btnStart.Text = "Start";
                timer1.Stop();
                camera.Stop();
                button9.Enabled = true;
                Text = "Pause";
            }

        }





        private static int inTimer = 0;
        private static int showWhat = 0; // 0: 顯示CCDImage ㄅ

        private int DrawCanvasHeight;


        private async void timer1_Tick(object sender, EventArgs e)
        {
            //在CCDIMAGE 上 畫出矩形
            Graphics g = CCDImage.CreateGraphics();
            Rectangle rect = new Rectangle(DrawCanvas.Location.X - CCDImage.Location.X, DrawCanvas.Location.Y, DrawCanvas.Width, DrawCanvas.Height);
            g.DrawRectangle(new Pen(Color.White, 2), rect);
            //--------------------------------------------------------------------------------------------------------------------------------------------
            //Text = camera.Size.Width.ToString() + "," + camera.Size.Height.ToString();
            //RotateNoneFlipX

            Bitmap myBitmap = camera.GetBitmap();
            myBitmap.RotateFlip(RotateFlipType.RotateNoneFlipX);

            scaleX = camera.Size.Width / CCDImage.Width;
            scaleY = camera.Size.Height / CCDImage.Height;

            int x = Convert.ToInt32(DrawCanvas.Left * scaleX);
            int y = Convert.ToInt32(DrawCanvas.Top * scaleY);
            int w = Convert.ToInt32(DrawCanvas.Width * scaleX);
            int h = Convert.ToInt32(DrawCanvasHeight * scaleY); //11/24
            Rectangle cloneRect = new Rectangle(x, y, w, h);
            System.Drawing.Imaging.PixelFormat format = myBitmap.PixelFormat;


            // Bitmap cloneBitmap = myBitmap.Clone(cloneRect, format);
            Bitmap cloneBitmap = crop(myBitmap, cloneRect);

            CCDImage.Image = myBitmap;
            ROIImage.Image = cloneBitmap;

            ROIImage.Left = DrawCanvas.Left;
            ROIImage.Top = DrawCanvas.Top;
            ROIImage.Height = DrawCanvas.Height;
            ROIImage.Width = DrawCanvas.Width;

            byte[] FULLBuffer = ImageToBuffer(CCDImage.Image, System.Drawing.Imaging.ImageFormat.Bmp);
            byte[] ROIBuffer = ImageToBuffer(ROIImage.Image, System.Drawing.Imaging.ImageFormat.Bmp);

            if (checkBox1.Checked) displayOriginal(ROIBuffer, x);
            if (checkBox9.Checked && isAutoScaling_END && isAvgEnd)
            {
                bool isClear = false;
                System.Windows.Forms.DataVisualization.Charting.Series seriesAVG = new System.Windows.Forms.DataVisualization.Charting.Series("Avg", 2000);


                //設定顏色
                seriesAVG.Color = Color.Black;

                //設定樣式
                seriesAVG.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                //迴圈二
                for (int i = 0; i < Original_Intensity.Count; i++)
                {
                    //給入數據畫圖
                    /* if (isClear)
                     {
                         this.chart_original.Series["seriesAVG"].Points.Clear();
                     }*/
                    this.chart_original.Series.Clear();
                    seriesAVG.Points.AddXY(i, Original_Intensity[i]);
                    this.chart_original.Series.Add(seriesAVG);
                    isClear = true;
                }
            }


            if (Interlocked.Exchange(ref inTimer, 1) == 1)
                return;

            if (ROIBuffer == null)
            {
                Interlocked.Exchange(ref inTimer, 0);
                return;
            }

            //if (checkBox1.Checked) displayOriginal(ROIBuffer, x);

            List<double> RealTime_Intensity = new List<double>(RealTime_Original_Intensity);

            //  Calibrate_EXP(FULLBuffer);
            Text = "Process...";
            /*
            if (Step == 1)
            { 
             var processTask1 = SpertroImageProcess(FULLBuffer, x);
            Task processFinishTask1 = await Task.WhenAny(processTask1);            
            }
            if (Step == 2)
            {
                var processTask2 = Peak_Looking_Task(ROIBuffer, Original_Intensity);
                Task processFinishTask2 = await Task.WhenAny(processTask2);
            }*/
            if (!Pause)
            {
                var processTask1 = SpertroImageProcess(FULLBuffer, x);
                Task processFinishTask1 = await Task.WhenAny(processTask1);
            }



            Interlocked.Exchange(ref inTimer, 0);
        }

        private Bitmap crop(Bitmap src, Rectangle cropRect)
        {
            // Rectangle cropRect = new Rectangle(0, 0, 400, 400);
            Bitmap target = new Bitmap(cropRect.Width, cropRect.Height);

            using (Graphics g = Graphics.FromImage(target))
            {
                g.DrawImage(src, new Rectangle(0, 0, target.Width, target.Height),
                      cropRect,
                      GraphicsUnit.Pixel);
            }
            return target;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            int cameraIndex = 0;
            // check format.
            string[] devices = UsbCamera.FindDevices();
            if (devices.Length == 0) return; // no camera.

            UsbCamera.VideoFormat[] formats = UsbCamera.GetVideoFormat(cameraIndex);
            //for (int i = 0; i < formats.Length; i++) Console.WriteLine("{0}:{1}", i, formats[i]);

            // create usb camera and start.
            camera = new UsbCamera(cameraIndex, formats[11]); //1280*960
            Console.WriteLine(formats[8] + "\n" + formats[9] + "\n" + formats[10] + "\n" + formats[11]);
        }

        int iDirectionLock = -1;
        private void DrawCanvas_MouseDown(object sender, MouseEventArgs e)
        {

            PictureBox pDrawROI = (PictureBox)sender;


            if (e.Button.Equals(MouseButtons.Left))
            {
                /*
                if (e.X < ROI_Edge)
                {
                    iDirectionLock = 1;
                }
                else */
                if (e.X > pDrawROI.Width - ROI_Edge)
                {
                    iDirectionLock = 2;
                }
                /*
                else if (e.Y < ROI_Edge)
                {
                    iDirectionLock = 3;
                }
                */
                else if (e.Y > pDrawROI.Height - ROI_Edge)
                {
                    iDirectionLock = 4;
                }
                else
                {
                    //移動 Top Left

                    iDirectionLock = 0;
                }
                ROI_Point = e.Location;
            }
        }


        private void DrawCanvas_MouseMove(object sender, MouseEventArgs e)
        {

            PictureBox pDrawROI = (PictureBox)sender;
            //bool bEdge = false;
            /*
            if (e.X < ROI_Edge)
            {
                pDrawROI.Cursor = Cursors.PanWest;
            }
            else */
            if (e.X > pDrawROI.Width - ROI_Edge)
            {
                pDrawROI.Cursor = Cursors.PanEast;
            }
            /*
            else if (e.Y < ROI_Edge)
            {
                pDrawROI.Cursor = Cursors.PanNorth;
            }
            */
            else if (e.Y > pDrawROI.Height - ROI_Edge)
            {
                pDrawROI.Cursor = Cursors.PanSouth;
            }
            else
            {
                pDrawROI.Cursor = Cursors.Cross;
            }
            int X = pDrawROI.Left;
            int Y = pDrawROI.Top;
            int W = pDrawROI.Width;
            int H = pDrawROI.Height;

            if (e.Button.Equals(MouseButtons.Left))
            {
                if (iDirectionLock == 1)
                {

                    X -= ROI_Point.X - e.Location.X;
                    W += ROI_Point.X - e.Location.X;

                }
                else if (iDirectionLock == 2)
                {
                    W += e.Location.X - ROI_Point.X;
                }

                else if (iDirectionLock == 3)
                {
                    X -= ROI_Point.Y - e.Location.Y;
                    H += ROI_Point.Y - e.Location.Y;

                }

                else if (iDirectionLock == 4)
                {
                    H += e.Location.Y - ROI_Point.Y;

                }

                if (iDirectionLock == 0)
                {
                    X += e.Location.X - ROI_Point.X;
                    Y += e.Location.Y - ROI_Point.Y;

                }
                else
                {
                    if (iDirectionLock == 1)
                    {

                        ROI_Point.Y = e.Location.Y;
                    }
                    else if (iDirectionLock == 2)
                    {
                        ROI_Point.X = e.Location.X;
                        ROI_Point.Y = e.Location.Y;
                    }
                    else if (iDirectionLock == 3)
                    {
                        ROI_Point.X = e.Location.X;

                    }
                    else if (iDirectionLock == 4)
                    {
                        ROI_Point.X = e.Location.X;
                        ROI_Point.Y = e.Location.Y;
                    }
                }

                if (X >= 0 && X + W <= CCDImage.Width)
                    pDrawROI.Left = X;
                if (Y >= 0 && Y + H <= CCDImage.Height)
                    pDrawROI.Top = Y;
                if (X + W >= 0 && X + W <= CCDImage.Width)
                    pDrawROI.Width = W;
                if (Y + H >= 0 && Y + H <= CCDImage.Height)
                    pDrawROI.Height = H;
            }
        }

        private void DrawCanvas_MouseUp(object sender, MouseEventArgs e)
        {
            iDirectionLock = -1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //---------------
            string html = File.ReadAllText("demo.html");
            webBrowser1.DocumentText = html;
            List<string> Scale = new List<string>();
            Scale = File.ReadAllLines(@"LaserScale\" + "AutoScale.txt").ToList();
            textBox8.Text = Scale[0];
            //----------------
            DrawCanvas.Visible = false;
            Get_SC_ID get_SC_ID = new Get_SC_ID();
            get_SC_ID.ShowDialog();
            if (get_SC_ID.button1.DialogResult.ToString()  == "OK")
            {
                //取得使用者輸入的SCID
                SC_ID.Text = get_SC_ID.textBox1.Text;
                //將SCID放入全域變數
                this_SC_ID = get_SC_ID.textBox1.Text;

                //--------------------------------------------------------------------------調整解析度
                if (get_SC_ID.checkBox1.Checked) //低解析度
                {
                    DrawCanvas.Parent = CCDImage;
                    ROIImage.Parent = panel1;
                    //初始化 ROI位置
                    ROI_Point.X = 0;
                    ROI_Point.Y = 0;
                    DrawCanvas.Left = ROI_Point.X;
                    DrawCanvas.Top = ROI_Point.Y;
                   
                    CCDImage.Width = 320;
                    CCDImage.Height = 240;
                    panel1.Width = 320;
                    panel1.Height = 240;
                    ROIImage.Width = 320;
                    ROIImage.Height = 240;

                    DrawCanvas.Width = CCDImage.Width;
                    DrawCanvas.Height = 5;
                    DrawCanvasHeight = 5;

                    panel1.Location = new Point(2, 252);
                    //ROIImage.Location = new Point(0, 250);

                }
                else if (get_SC_ID.checkBox2.Checked) //高解析度
                {
                    DrawCanvas.Parent = CCDImage;
                    ROIImage.Parent = panel1;
                    //初始化 ROI位置
                    ROI_Point.X = 0;
                    ROI_Point.Y = 0;
                    DrawCanvas.Left = ROI_Point.X;
                    DrawCanvas.Top = ROI_Point.Y;

                    CCDImage.Width = 640;
                    CCDImage.Height = 480;
                    panel1.Width = 640;
                    panel1.Height = 480;
                    ROIImage.Width = 640;
                    ROIImage.Height = 480;

                    DrawCanvas.Width = CCDImage.Width;
                    DrawCanvas.Height = 20 / 2;
                }
                //--------------------------------讀入雷射標準FWHM---------------------
                LockBox.Checked = true;
                List<string> LaserScale = new List<string>();

                LaserScale = File.ReadAllLines(@"LaserScale\" + "LaserScale.txt").ToList();
                L1.Text = LaserScale[0];
                L2.Text = LaserScale[1];
                L3.Text = LaserScale[2];
                L4.Text = LaserScale[3];
                L5.Text = LaserScale[4];
                L6.Text = LaserScale[5];
                L7.Text = LaserScale[6];
                L8.Text = LaserScale[7];
                //--------------------------------讀入汞氬燈標準FWHM----------------------
                LockBox.Checked = true;
                List<string> HgScale = new List<string>();

                LaserScale = File.ReadAllLines(@"LaserScale\" + "HgScale.txt").ToList();
                HG1.Text = LaserScale[0];
                HG2.Text = LaserScale[1];
                HG3.Text = LaserScale[2];
                HG4.Text = LaserScale[3];
                HG5.Text = LaserScale[4];
                HG6.Text = LaserScale[5];
                //-------------------------------------------------------------------------

                //Load_Agri();

                //連接 Console
                SerialPort serialPort1 = new SerialPort();
                string[] portnames = SerialPort.GetPortNames();
                foreach (var item in portnames)
                {
                    combobox1.Items.Add(item);
                }

                Console_Connect(combobox1.Text, Convert.ToInt32(TB_Buad.Text));
                // Console_Input.Text = "尚未連結COM";

                checkBox1.Checked = true;

                try
                {
                    Directory.Delete("C://Users//" + Environment.UserName + "//AppData//Local//Temp//" + Environment.UserName + "//mcrCache9.4", true);
                }
                catch (Exception)
                {
                    Console.WriteLine("路徑不存在");
                }
            }

        }

        //Smart_Scaling






        private void btnROI_Click(object sender, EventArgs e)
        {
            // progressROI.Value = 0;





        }
        //------------------------------序列埠------------------------------------------------//

        //連接 Console
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

                    Console_receiving = true;

                    //開啟執行續做接收動作
                    //t = new Thread(DoReceive);

                    //t.IsBackground = true;
                    //t.Start();
                    Console_Input.Text = "";
                    Console_Input.Text = "連結成功";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //關閉 Console
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

        //Console 接收資料
        private string DoReceive()
        {
            Byte[] buffer = new Byte[1024];
            string buf = "";
            try
            {
                while (Console_receiving)
                {
                    if (My_SerialPort.BytesToRead > 0)
                    {
                        Int32 length = My_SerialPort.Read(buffer, 0, buffer.Length);

                        //string buf = Encoding.ASCII.GetString(buffer);
                        buf = Encoding.ASCII.GetString(buffer);
                        Array.Resize(ref buffer, length);
                        Display d = new Display(ConsoleShow);
                        this.Invoke(d, new Object[] { buf });
                        Array.Resize(ref buffer, 1024);
                    }

                    Thread.Sleep(20);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return buf;

        }
        public void ConsoleShow(string buffer)
        {
            Console_Input.AppendText(buffer);
        }

        //Console 發送資料
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

        private void btn_Send_Click(object sender, EventArgs e)
        {
            SendData(Encoding.ASCII.GetBytes(Console_Output.Text + "\r\n"));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SendData(Encoding.ASCII.GetBytes("LAS:OUT?" + "\r\n"));
        }

        private void Connectbutton_Click(object sender, EventArgs e)
        {
            //連接 Console
            SerialPort serialPort1 = new SerialPort();
            String[] portnames = SerialPort.GetPortNames();
            foreach (var item in portnames)
            {
                combobox1.Items.Add(item);
            }

            Console_Connect(combobox1.Text, Convert.ToInt32(TB_Buad.Text));
        }

        //-------------------------------------Display---------------------//
        private void displayOriginal(byte[] input_image_buffer, int start_pixel)//育代
        {
            //this.chart2.Series.Clear();
            Bitmap input_image = BufferToImage(input_image_buffer);
            int W;
            int H;
            W = input_image.Width;
            H = input_image.Height;
            //Bitmap image_roi_for_gray = new Bitmap(w, h);

            int Pixel_x = 0;//正在被掃描的點
            int Pixel_y = 0;


            //var sgf = MathNet.Filtering.OnlineFilter.CreateDenoise(window);


            System.Windows.Forms.DataVisualization.Charting.Series seriesGray = new System.Windows.Forms.DataVisualization.Charting.Series("灰階", 2000);

            System.Windows.Forms.DataVisualization.Charting.Series seriesSG1 = new System.Windows.Forms.DataVisualization.Charting.Series("SG", 2000);

            System.Windows.Forms.DataVisualization.Charting.Series seriesCali = new System.Windows.Forms.DataVisualization.Charting.Series("校正", 2000);

            RealTime_Original_Intensity = Math_Methods.get_Original_Intensity(input_image);

            SG_Intensity = RealTime_Original_Intensity;





            //設定座標大小
            this.chart_original.ChartAreas[0].AxisY.Minimum = 0;
            this.chart_original.ChartAreas[0].AxisY.Maximum = 300;

            input_image.Dispose();

            //設定標題

            this.chart_original.Titles.Clear();
            this.chart_original.Titles.Add("S01");
            if (checkBox6.Checked)
            {
                this.chart_original.Titles[0].Text = "原始光譜_" + this_SC_ID + "_Hg_Ar";
            }
            else if (checkBox7.Checked)
            {
                this.chart_original.Titles[0].Text = "原始光譜_" + this_SC_ID + "_Laser";
            }
            else if (checkBox8.Checked)
            {
                this.chart_original.Titles[0].Text = "原始光譜_" + this_SC_ID + "_White";
            }
            this.chart_original.Titles[0].ForeColor = Color.Black;
            this.chart_original.Titles[0].Font = new System.Drawing.Font("標楷體", 16F);
            // this.chart_original.


            //設定顏色
            seriesGray.Color = Color.Blue;
            seriesSG1.Color = Color.Orange;
            seriesCali.Color = Color.Red;
            //設定樣式
            seriesGray.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            seriesSG1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            //迴圈二
            for (Pixel_x = 0; Pixel_x < W; Pixel_x++)
            {
                //給入數據畫圖
                seriesGray.Points.AddXY(Pixel_x + start_pixel, RealTime_Original_Intensity[Pixel_x]);
                seriesSG1.Points.AddXY(Pixel_x + start_pixel, SG_Intensity[Pixel_x]);
                this.chart_original.Series.Clear();

                if (checkBox1.Checked)
                    this.chart_original.Series.Add(seriesGray);
                if (checkBox2.Checked)
                    this.chart_original.Series.Add(seriesSG1);
            }


        }


        /*AutoScalling*/
        private int now_exp = 0;
        private bool Calibrate_EXP(byte[] Input_Image)
        {
            //  int prop = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            int speed;
            int Max;
            int def_exp = -5;
            int max_exp = -2;
            int min_exp = -11;
            //  displayOriginal(imgbuffer, x, x + imgbuffer.Width);
            //影像處理程序 把Do While 刪除 寫在這
            //測試用

            //     Test_Image = Input_Image.Clone(cloneRect_original, format);

            List<int> Speed_and_Max = new List<int>();
            Speed_and_Max = Math_Methods.Control_EXP(Input_Image);
            speed = Speed_and_Max[0];
            Max = Speed_and_Max[1];
            //    var Last_EXP = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            UsbCamera.PropertyItems.Property prop;
            prop = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            if (prop.Available)
            {
                max_exp = prop.Max;
                min_exp = prop.Min;
                def_exp = prop.Default;

            }
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            switch (speed)
            {
                case 2:
                    // this.Invoke(camera_prop, "exp", (def_exp - min_exp) /2+ min_exp);
                    this.Invoke(camera_prop, "exp", def_exp - 1);
                    break;
                case 1:
                    this.Invoke(camera_prop, "exp", def_exp - 1);
                    break;
                case 0:
                    Console.WriteLine("Good");
                    break;
                case -1:
                    this.Invoke(camera_prop, "exp", def_exp + 1);
                    break;
                case -2:
                    this.Invoke(camera_prop, "exp", def_exp + 1);
                    break;
                case -3:
                    this.Invoke(camera_prop, "exp", def_exp + 1);
                    break;
                default:
                    Console.WriteLine("The color is unknown.");
                    break;
            }




            if (speed == 0) return true;
            else return false;

        }

        private List<double> Smart_Calibrate_DG_Intensity(List<double> Input_List, double index) //.
        {




            int Max = 1;
            int def_dg = 100;
            int max_dg = 255;
            int min_dg = 32;
            List<double> result = new List<double>();



            UsbCamera.PropertyItems.Property prop;
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
            if (prop.Available)
            {
                max_dg = prop.Max;
                min_dg = prop.Min;
                def_dg = prop.Default;

            }


            //  this.Invoke(camera_prop, "gain", def_dg);
            if (index == 0)
            {
                double Max_Intensity = Input_List.Max();
                int Index_of_Max = Input_List.IndexOf(Max_Intensity);    
                result.Add(def_dg);//[0] 是否過曝  0:
                result.Add(Max_Intensity);//[0] 此時的值
                result.Add(Index_of_Max);//[1] 發生的位置
                this.Invoke(camera_prop, "gain", Convert.ToInt32(def_dg / 2));
            }
            else
            {
                double New_Max_Intensity = Input_List[Convert.ToInt32(index)];
                result.Add(def_dg);//[0] 是否過曝  0:
                result.Add(New_Max_Intensity);//[0] 此時的值
                result.Add(Convert.ToInt32(index));//[1] 發生的位置


            }



            return result;


        }

        private bool Calibrate_DG_Intensity(List<double> Input_List)
        {
            //  int prop = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            double speed;
            double Max;
            int def_dg = -5;
            int max_dg = -2;
            int min_dg = -11;
            //  displayOriginal(imgbuffer, x, x + imgbuffer.Width);
            //影像處理程序 把Do While 刪除 寫在這
            //測試用

            //     Test_Image = Input_Image.Clone(cloneRect_original, format);

            List<double> Speed_and_Max = new List<double>();
            Speed_and_Max = Math_Methods.Control_EXP_List(Input_List);
            speed = Speed_and_Max[0];
            Max = Speed_and_Max[1];
            //    var Last_EXP = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            UsbCamera.PropertyItems.Property prop;
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
            if (prop.Available)
            {
                max_dg = prop.Max;
                min_dg = prop.Min;
                def_dg = prop.Default;

            }
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            switch (speed)
            {
                case 2:
                    // this.Invoke(camera_prop, "exp", (def_exp - min_exp) /2+ min_exp);
                    this.Invoke(camera_prop, "gain", def_dg - 25);
                    break;
                case 1:
                    this.Invoke(camera_prop, "gain", def_dg - 8);
                    break;
                case 0:
                    Console.WriteLine("Good");
                    break;
                case -1:
                    this.Invoke(camera_prop, "gain", def_dg + 5);
                    break;
                case -2:
                    this.Invoke(camera_prop, "gain", def_dg + 30);
                    break;
                case -3:
                    this.Invoke(camera_prop, "gain", def_dg + 50);
                    break;
                default:
                    Console.WriteLine("The color is unknown.");
                    break;
            }

            if (speed == 0) return true;
            else return false;

        }
        private bool Calibrate_DG(byte[] Input_Image)
        {
            //  int prop = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            int speed;
            int Max;
            int def_dg = -5;
            int max_dg = -2;
            int min_dg = -11;
            //  displayOriginal(imgbuffer, x, x + imgbuffer.Width);
            //影像處理程序 把Do While 刪除 寫在這
            //測試用

            //     Test_Image = Input_Image.Clone(cloneRect_original, format);

            List<int> Speed_and_Max = new List<int>();
            Speed_and_Max = Math_Methods.Control_EXP(Input_Image);
            speed = Speed_and_Max[0];
            Max = Speed_and_Max[1];
            //    var Last_EXP = camera.Properties[DirectShow.CameraControlProperty.Exposure];
            UsbCamera.PropertyItems.Property prop;
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
            if (prop.Available)
            {
                max_dg = prop.Max;
                min_dg = prop.Min;
                def_dg = prop.Default;

            }
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            switch (speed)
            {
                case 2:
                    // this.Invoke(camera_prop, "exp", (def_exp - min_exp) /2+ min_exp);
                    this.Invoke(camera_prop, "gain", def_dg - 25);
                    break;
                case 1:
                    this.Invoke(camera_prop, "gain", def_dg - 8);
                    break;
                case 0:
                    Console.WriteLine("Good");
                    break;
                case -1:
                    this.Invoke(camera_prop, "gain", def_dg + 5);
                    break;
                case -2:
                    this.Invoke(camera_prop, "gain", def_dg + 20);
                    break;
                case -3:
                    this.Invoke(camera_prop, "gain", def_dg + 50);
                    break;
                default:
                    Console.WriteLine("The color is unknown.");
                    break;
            }

            if (speed == 0) return true;
            else return false;

        }

        private void Set_Dg_half() {

            int def_dg = 100;
            int max_dg = 255;
            int min_dg = 32;

            UsbCamera.PropertyItems.Property prop;
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
            if (prop.Available)
            {
                max_dg = prop.Max;
                min_dg = prop.Min;
                def_dg = prop.Default;

            }

            this.Invoke(camera_prop, "gain", Convert.ToInt32(def_dg));
        }

        private void Set_Dg(int new_dg)
        {
            Dg_Update dg_Update = new Dg_Update(Update_Dg);
            this.Invoke(dg_Update, new_dg);
            int def_dg = 100;
            int max_dg = 255;
            int min_dg = 32;

            UsbCamera.PropertyItems.Property prop;
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
            if (prop.Available)
            {
                max_dg = prop.Max;
                min_dg = prop.Min;
                def_dg = prop.Default;

            }
            if (new_dg >= 255) {
                this.Invoke(camera_prop, "gain", 255);
               // MessageBox.Show("其他參數調太暗了!已無法更亮");
            } else if (new_dg <= 32) {
                this.Invoke(camera_prop, "gain", 32);
                //MessageBox.Show("其他參數調太亮了!已無法更暗");

            }
            else {
               this.Invoke(camera_prop, "gain", new_dg);
            }
           
        }

       private void Set_Gamma(int new_gamma)
        {
            Gamma_Update gamma_Update = new Gamma_Update(Update_Gamma);
            this.Invoke(gamma_Update, new_gamma);
            int def_gamma = 100;
            int max_gamma = 100;
            int min_gamma = 300;

            UsbCamera.PropertyItems.Property prop;
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gamma];
            if (prop.Available)
            {
                max_gamma = prop.Max;
                min_gamma = prop.Min;
                def_gamma = prop.Default;

            }
            if (new_gamma >= 300)
            {
                this.Invoke(camera_prop, "gamma", 300);
                //MessageBox.Show("其他參數調太暗了!已無法更亮");
            }
            else if (new_gamma <= 100)
            {
                this.Invoke(camera_prop, "gamma", 100);
                //MessageBox.Show("其他參數調太亮了!已無法更暗");

            }
            else
            {
                this.Invoke(camera_prop, "gamma", new_gamma);
            }

        }

        private void Set_Back(int new_back)
        {
            Gamma_Update back_Update = new Gamma_Update(Update_Back);
            this.Invoke(back_Update, new_back);
            int def_back = 1;
            int max_back = 3;
            int min_back = 0;

            UsbCamera.PropertyItems.Property prop;
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.BacklightCompensation];
            if (prop.Available)
            {
                max_back = prop.Max;
                min_back = prop.Min;
                def_back = prop.Default;

            }
            if (new_back >= 3)
            {
                this.Invoke(camera_prop, "back", 3);
                //MessageBox.Show("其他參數調太暗了!已無法更亮");
            }
            else if (new_back <= 0)
            {
                this.Invoke(camera_prop, "back", 0);
                //MessageBox.Show("其他參數調太亮了!已無法更暗");

            }
            else
            {
                this.Invoke(camera_prop, "back", new_back);
            }

        }


        private List<double> Smart_Calibrate_DG(byte[] Input_Image , double index) //.
        {




            int Max = 1;
            int def_dg = 100;
            int max_dg = 255;
            int min_dg = 32;
            List<double> result = new List<double>();



            UsbCamera.PropertyItems.Property prop;
            set_camera_prop_dele camera_prop = new set_camera_prop_dele(set_camera_prop);
            prop = camera.Properties[DirectShow.VideoProcAmpProperty.Gain];
            if (prop.Available)
            {
                max_dg = prop.Max;
                min_dg = prop.Min;
                def_dg = prop.Default;

            }


            //  this.Invoke(camera_prop, "gain", def_dg);
            if (index == 0)
            {
                int Max_Intensity = Input_Image.Max();
                int Index_of_Max = Array.IndexOf(Input_Image, Max_Intensity);
                result.Add(def_dg);//[0] 是否過曝  0:
                result.Add(Max_Intensity);//[0] 此時的值
                result.Add(Index_of_Max);//[1] 發生的位置
                this.Invoke(camera_prop, "gain",Convert.ToInt32( def_dg/2));
            }
            else
            {
                int New_Max_Intensity = Input_Image[Convert.ToInt32(index)];
                result.Add(def_dg);//[0] 是否過曝  0:
                result.Add(New_Max_Intensity);//[0] 此時的值
                result.Add(Convert.ToInt32(index));//[1] 發生的位置


            }
            
           

            return result;


        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }



        //當load進來時，先用CHAN?確認當前Laser，並寫至updn_manual
        private void updn_manual_ValueChanged(object sender, EventArgs e)
        {
            //手動雷射控制//
            /*
             1.設定當前雷射 CHAN 1
             2.確認當前雷射狀態LAS:OUT?
             3. 更改btn_Laser.text
             4. 當前雷射溫度、電流(之後會加timer來讀取)
             
             */



        }
        /*
        private async Task<bool> OnOff_Laser(bool OnOff) //完成後回傳TRUE
        {
            //await Thread.Sleep(2000);
            //await Task.Delay(3000);

            int i = 0;
        

            //Task.Run(() => ShowThreadInfo("Task"));
            await Task.Run(() =>
            {
                //影像處理程序 把Do While 刪除 寫在這
                //測試用
                do
                {
                    i++;
                    Text = i.ToString();
                    FormUpdata formcontrl = new FormUpdata(FormUpdataMethod);
                    this.Invoke(formcontrl, i);

                } while (i < 1000000000);

            });

            return true;
        }
        */

        private void btn_Laser_Click(object sender, EventArgs e)
        {
            switch (btn_Laser.Text)
            {
                case "ON":
                    SendData(Encoding.ASCII.GetBytes("LAS:OUT 0" + "\r\n"));

                    break;
                case "OFF":
                    SendData(Encoding.ASCII.GetBytes("LAS:OUT 1" + "\r\n"));

                    break;
                default:
                    Console.WriteLine("????");
                    break;
            }
            Task.Delay(1000);



        }

        private bool Laser_On(int Channel)
        {
            //切換頻道
            SendData(Encoding.ASCII.GetBytes("channel " + Channel + "\r\n"));
            Task.Delay(1000);
            //確認該雷射狀態
            SendData(Encoding.ASCII.GetBytes("LAS:OUT?" + "\r\n"));
            Task.Delay(1000);




            return true;
        }


        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {

        }

       

        private void btn_reset_Click(object sender, EventArgs e)
        {
            initTask();
        }

        private void btn_Ch1_Click(object sender, EventArgs e)
        {
            isPeakLookingFinish = true;
            pFWHM[0] = Gaus_Pixel_Intensity_Parameter_Set["Parameters"][3];
            Console.WriteLine(pFWHM);
        }

        private void btn_MatlabLoad_Click(object sender, EventArgs e)
        {
            Load_Agri();
            MessageBox.Show("載入成功");
        }

        void Load_Agri() {

            double[] TestArray = new double[100];
            //     
            //    List<double> Test_Intensity = new List<double>(1280);
            // Instantiate random number generator using system-supplied value as seed.
            var rand = new Random();

            for (int ctr = 0; ctr < 100; ctr++)
            {
                //   Console.Write("{0,8:N0}", rand.Next(15));
                if (ctr > 50 && ctr < 60)
                {
                    TestArray[ctr] = TestArray[ctr] + rand.NextDouble() * 10 + 200;
                }
                else
                {
                    TestArray[ctr] = TestArray[ctr] + rand.NextDouble() * 10;
                }




            }
            Console.WriteLine();
            List<double> Test_Intensity = TestArray.ToList<double>();


            Math_Methods.SG_Fitting(Test_Intensity, 3, 21);
            //Math_Methods.Auto_Gaussian(Math_Methods.SG_Fitting(Test_Intensity, 3, 21));
            // Math_Methods.Auto_Gaussian(Test_Intensity);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 0;CHAN 2;LAS:OUT 0;CHAN 3;LAS:OUT 0;CHAN 4;LAS:OUT 0;CHAN 5;LAS:OUT 0;CHAN 6;LAS:OUT 0;CHAN 7;LAS:OUT 0;CHAN 8;LAS:OUT 0;\r\n"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 1;CHAN 2;LAS:OUT 1;CHAN 3;LAS:OUT 1;CHAN 4;LAS:OUT 1;CHAN 5;LAS:OUT 1;CHAN 6;LAS:OUT 1;CHAN 7;LAS:OUT 1;CHAN 8;LAS:OUT 1;\r\n"));
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

       

        private void button6_Click(object sender, EventArgs e)
        {

            if (Pause)
            {
                Pause = false;
                button6.Text = "暫停";
            }
            else { 
            Pause = true;
                button6.Text="繼續";
            }
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }


        private void button9_Click(object sender, EventArgs e)
        {
            SendData(Encoding.ASCII.GetBytes("CHAN 1;LAS:OUT 1;CHAN 2;LAS:OUT 1;CHAN 3;LAS:OUT 1;CHAN 4;LAS:OUT 1;CHAN 5;LAS:OUT 1;CHAN 6;LAS:OUT 1;CHAN 7;LAS:OUT 1;CHAN 8;LAS:OUT 1;\r\n"));
            Form_camp FP = new Form_camp(Camera_Parameters["gamma"], Camera_Parameters["back"]);
            this.Enabled = false;
            this.Visible = false;
            FP.Show();
            Console_Connect(combobox1.Text, Convert.ToInt32(TB_Buad.Text));
            Thread.Sleep(1000);

        }

       

        private void btn_Start1_Click_1(object sender, EventArgs e)
        {
            isFirtCacular++;
            Pause = false;
            iTask = 1;
            isHg_Ar = false;
            isLaser = true;
            checkBox7.Checked = true;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            /* JObject videogameRatings = new JObject(
     new JProperty("Halo", 9),
     new JProperty("Starcraft", 9),
     new JProperty("Call of Duty", 7.5));*/

            File.WriteAllText(@"E:\isb.json", richTextBox2.Text);
        }

        private void btn_Start2_Click_1(object sender, EventArgs e)
        {
            textBox13.Text = "40";
            isLaser = true;
            Pause = false;
            isHg_Ar = false;
            iTask = 400;
            checkBox7.Checked = true;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            iTask = 1000;
            Pause = false;
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            if (Directory.Exists(@"" + tb_path0.Text + tb_path1.Text))
            {
                Console.WriteLine("The directory {0} already exists.", @"" + tb_path0.Text + tb_path1.Text);
            }
            else
            {
                Directory.CreateDirectory(@"" + tb_path0.Text + tb_path1.Text);
                Console.WriteLine("The directory {0} was created.", @"" + tb_path0.Text + tb_path1.Text);
            }
            if (Directory.Exists(@"" + tb_path0.Text + tb_path2.Text))
            {
                Console.WriteLine("The directory {0} already exists.", @"" + tb_path0.Text + tb_path2.Text);
            }
            else
            {
                Directory.CreateDirectory(@"" + tb_path0.Text + tb_path2.Text);
                Console.WriteLine("The directory {0} was created.", @"" + tb_path0.Text + tb_path2.Text);
            }
            if (Directory.Exists(@"" + tb_path0.Text + tb_path3.Text))
            {
                Console.WriteLine("The directory {0} already exists.", @"" + tb_path0.Text + tb_path3.Text);
            }
            else
            {
                Directory.CreateDirectory(@"" + tb_path0.Text + tb_path3.Text);
                Console.WriteLine("The directory {0} was created.", @"" + tb_path0.Text + tb_path3.Text);
            }
            if (Directory.Exists(@"" + tb_path0.Text + tb_path4.Text))
            {
                Console.WriteLine("The directory {0} already exists.", @"" + tb_path0.Text + tb_path4.Text);
            }
            else
            {
                Directory.CreateDirectory(@"" + tb_path0.Text + tb_path4.Text);
                Console.WriteLine("The directory {0} was created.", @"" + tb_path0.Text + tb_path4.Text);
            }



            // File.WriteAllText(@"E:\isb.json", richTextBox1.Text);
            //    File.WriteAllText(@"E:\isb.json", richTextBox3.Text);
            File.WriteAllText(@"" + tb_path0.Text + tb_path1.Text + "/" + textBox2.Text + "_pixel-wave.txt", richTextBox1.Text);
            File.WriteAllText(@"" + tb_path0.Text + tb_path2.Text + "/" + textBox2.Text + "_para.txt", richTextBox3.Text);
            File.WriteAllText(@"" + tb_path0.Text + tb_path3.Text + "/" + textBox2.Text + "_isb.json", richTextBox2.Text);
            List<string> original_int = new List<string>();
            for (int i = 0; i < Pure_Intensity.Count; i++)
            {
                original_int.Add(Pure_Intensity[i].ToString());
                
            }//@"" + tb_path0.Text + tb_path5.Text + "/" + textBox2.Text +
            File.WriteAllLines(textBox2.Text + "_original_int.txt", original_int.ToArray());
            MessageBox.Show("存檔成功");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /* List<string> original_int = new List<string>();
             for (int i = 0; i < Pure_Intensity.Count; i++)
             {
                 original_int.Add(Pure_Intensity[i].ToString());
             }//@"" + tb_path0.Text + tb_path5.Text + "/" + textBox2.Text +
             File.WriteAllLines(textBox2.Text+"_original_int.txt", original_int.ToArray());
             MessageBox.Show("存檔成功");*/         
            
        }
        #region 視窗截圖
        /* 引入 Win32 API 中的 User32.DLL
         * 需要加上 using System.Runtime.InteropServices;
        */
        [DllImport("user32.dll")]
        public static extern Boolean GetWindowRect(IntPtr hWnd, ref CapRectangle bounds);

        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int Width, int Height, int flags);
        /// <summary>
        /// 範圍
        /// </summary>
        //[StructLayout(LayoutKind.Sequential)] 不知道幹嘛的
        public struct CapRectangle
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion
        private void btn_Ch2_Click(object sender, EventArgs e)
        {
            isch2 = true;
            btn_Ch2.BackColor = Color.Green;
        }

        private void btn_Ch3_Click(object sender, EventArgs e)
        {
            isch3 = true;
            btn_Ch3.BackColor = Color.Green;
        }

        private void btn_Ch4_Click(object sender, EventArgs e)
        {
            isch4 = true;
            btn_Ch4.BackColor = Color.Green;
        }

        private void btn_Ch5_Click(object sender, EventArgs e)
        {
            isch5 = true;
            btn_Ch5.BackColor = Color.Green;
        }

        private void btn_Ch6_Click(object sender, EventArgs e)
        {
            isch6 = true;
            btn_Ch6.BackColor = Color.Green;
        }

        private void btn_Ch7_Click(object sender, EventArgs e)
        {
            isch7 = true;
            btn_Ch7.BackColor = Color.Green;
        }

        private void btn_Ch8_Click(object sender, EventArgs e)
        {
            isch8 = true;
            btn_Ch8.BackColor = Color.Green;
        }

        private void btn_laser_reset_Click(object sender, EventArgs e)
        {
            isch1 = false; isch2 = false; isch3 = false; isch4 = false; isch5 = false; isch6 = false; isch7 = false; isch8 = false;
            btn_Ch1.BackColor = Color.LightGray;
            btn_Ch2.BackColor = Color.LightGray;
            btn_Ch3.BackColor = Color.LightGray;
            btn_Ch4.BackColor = Color.LightGray;
            btn_Ch5.BackColor = Color.LightGray;
            btn_Ch6.BackColor = Color.LightGray;
            btn_Ch7.BackColor = Color.LightGray;
            btn_Ch8.BackColor = Color.LightGray;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            List<string> original_int = new List<string>();
            for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
            {
                original_int.Add(RealTime_Original_Intensity[i].ToString());

            }//@"" + tb_path0.Text + tb_path5.Text + "/" + textBox2.Text +

            Directory.CreateDirectory(@"Screenshot\" + this_SC_ID);
            chart_original.SaveImage(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg", ImageFormat.Bmp);

            Directory.CreateDirectory(@"RawData\Manual_Control\" + this_SC_ID);
            File.WriteAllLines(@"RawData\Manual_Control\" + this_SC_ID + @"\" + "Intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", original_int.ToArray());
            MessageBox.Show("存檔成功");
            

        }

        private void button12_Click(object sender, EventArgs e)
        {
            json js = new json(this);
            js.Show();
        }

        private void LockBox_CheckedChanged(object sender, EventArgs e)
        {
            if (LockBox.Checked)
            {
                List<string> LaserScale = new List<string>();
                LaserScale.Add(L1.Text);
                LaserScale.Add(L2.Text);
                LaserScale.Add(L3.Text);
                LaserScale.Add(L4.Text);
                LaserScale.Add(L5.Text);
                LaserScale.Add(L6.Text);
                LaserScale.Add(L7.Text);
                LaserScale.Add(L8.Text);
                File.WriteAllLines(@"LaserScale\" + "LaserScale.txt", LaserScale);
     
                L1.Text = LaserScale[0];
                L1.Enabled = false;
                L2.Enabled = false;
                L3.Enabled = false;
                L4.Enabled = false;
                L5.Enabled = false;
                L6.Enabled = false;
                L7.Enabled = false;
                L8.Enabled = false;
            }
            else if (LockBox.Checked == false)
            {
                L1.Enabled = true;
                L2.Enabled = true;
                L3.Enabled = true;
                L4.Enabled = true;
                L5.Enabled = true;
                L6.Enabled = true;
                L7.Enabled = true;
                L8.Enabled = true;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            List<string> LaserScale = new List<string>();

            LaserScale = File.ReadAllLines(@"LaserScale\" + "LaserScale_Original.txt").ToList();
            L1.Text = LaserScale[0];
            L2.Text = LaserScale[1];
            L3.Text = LaserScale[2];
            L4.Text = LaserScale[3];
            L5.Text = LaserScale[4];
            L6.Text = LaserScale[5];
            L7.Text = LaserScale[6];
            L8.Text = LaserScale[7];
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            Set_Dg(trackBar1.Value);
            label29.Text = trackBar1.Value.ToString();
        }

        private void button14_Click(object sender, EventArgs e)
        { 
            if (textBox13.Text == "40")
            {
                textBox13.Text = "20";
            }
            Pause = false;
            iTask = 625;
            isHg_Ar = true;
            isFirtCacular++;
            btn_Start1.Enabled = true;
            btn_Start2.Enabled = true;
            Hg2WFHM = true;
            Hg3WFHM = true;
            checkBox6.Checked = true;
        }

        private void LockBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (LockBox2.Checked)
            {
                List<string> HgScale = new List<string>();
                HgScale.Add(HG1.Text);
                HgScale.Add(HG2.Text);
                HgScale.Add(HG3.Text);
                HgScale.Add(HG4.Text);
                HgScale.Add(HG5.Text);
                HgScale.Add(HG6.Text);
             
                File.WriteAllLines(@"LaserScale\" + "HgScale.txt", HgScale);
                HG1.Enabled = false;
                HG2.Enabled = false;
                HG3.Enabled = false;
                HG4.Enabled = false;
                HG5.Enabled = false;
                HG6.Enabled = false;
            }
            else if (LockBox2.Checked == false)
            {
                HG1.Enabled = true;
                HG2.Enabled = true;
                HG3.Enabled = true;
                HG4.Enabled = true;
                HG5.Enabled = true;
                HG6.Enabled = true;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            List<string> HgScale = new List<string>();
            HgScale = File.ReadAllLines(@"LaserScale\" + "HgScale_Original.txt").ToList();
            HG1.Text = HgScale[0];
            HG2.Text = HgScale[1];
            HG3.Text = HgScale[2];
            HG4.Text = HgScale[3];
            HG5.Text = HgScale[4];
            HG6.Text = HgScale[5];
           
        }

        private void Laser_And_Hg_Ar_CHK_CheckedChanged(object sender, EventArgs e)
        {
            if (Laser_And_Hg_Ar_CHK.Checked)
            {
                btn_Start1.Enabled = false;
                btn_Start2.Enabled = false;
            }
            else {
                btn_Start1.Enabled = true;
                btn_Start2.Enabled = true;
            }
        }

        private Bitmap Overlay_Rgb(Bitmap p1, Bitmap p2)
        {
            Bitmap p3 = new Bitmap(p1.Width, p1.Height);
            Color R, G, B;
            for (int i = 0; i < p1.Width; i++)
            {
                for (int j = 0; j < p2.Height; j++)
                {
                    Color c1 = p1.GetPixel(i, j);
                    Color c2 = p2.GetPixel(i, j);
                    Color c3 = Color.FromArgb(0, 0, 0);
                    int r, g, b;
                    if (c1.R + c2.R >= 255)
                    {
                        r = 255;
                        if (c1.G + c2.G >= 255)
                        {
                            g = 255;
                            if (c1.B + c2.B >= 255)
                            {
                                b = 255;
                            }
                            else
                            {
                                b = c1.B + c2.B;
                            }

                        }
                        else
                        {
                            g = c1.G + c2.G;
                            if (c1.B + c2.B >= 255)
                            {
                                b = 255;
                            }
                            else
                            {
                                b = c1.B + c2.B;
                            }
                        }
                    }
                    else
                    {
                        r = c1.R + c2.R;
                        if (c1.G + c2.G >= 255)
                        {
                            g = 255;
                            if (c1.B + c2.B >= 255)
                            {
                                b = 255;
                            }
                            else
                            {
                                b = c1.B + c2.B;
                            }

                        }
                        else
                        {
                            g = c1.G + c2.G;
                            if (c1.B + c2.B >= 255)
                            {
                                b = 255;
                            }
                            else
                            {
                                b = c1.B + c2.B;
                            }
                        }
                    }

                    c3 = Color.FromArgb(r, g, b);
                    p3.SetPixel(i, j, c3);
                }
            }
            return p3;
        }

        private void chartScreenshot(string name)
        {
            
            chart_original.SaveImage(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg", ImageFormat.Bmp);
            CCDImage.Image.Save(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID+"_Image_"+ $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg");
            ROIImage.Image.Save(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_ROI_" + $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg");

            if (checkBox5.Checked && isGetDarkChart)
            {
                chart1.Titles["單根雷射疊圖"].Text = this_SC_ID + "單根雷射疊圖";
                switch (Laser)
                {
                    case 1:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series1"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        break;
                    case 2:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series2"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        break;
                    case 3:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series3"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        break;
                    case 5:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series4"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        break;
                    case 6:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series5"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        break;
                    case 7:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series6"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        break;
                    case 8:
                        for (int i = 0; i < RealTime_Original_Intensity.Count; i++)
                        {
                            chart1.Series["Series7"].Points.AddXY(i, RealTime_Original_Intensity[i]);
                        }
                        chart1.SaveImage(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_SingleLaser_Overlay_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg", ImageFormat.Bmp);
                        break;
                }
            }
            //--------------------------------------截圖含roi的image-----------------------------------------------------------------------------------
            if (checkBox5.Checked && isHg_Ar == false && onlyWhite == false && Laser == 1)
            {
                Bitmap myimage = new Bitmap(this.CCDImage.Width + 5, this.CCDImage.Height + 10);
                Graphics gg = Graphics.FromImage(myimage);
                gg.CopyFromScreen(new Point(this.Location.X + CCDImage.Location.X + 5, this.Location.Y + CCDImage.Location.Y + 28), new Point(CCDImage.Location.X, CCDImage.Location.Y), new Size(CCDImage.Width + 15, CCDImage.Height + 20));
                IntPtr dc1 = gg.GetHdc();
                gg.ReleaseHdc(dc1);
                myimage.Save(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_full_Image_" + $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg");
                p1 = myimage;
            }
            else if (checkBox5.Checked && isHg_Ar == false && onlyWhite == false && Laser != 1)
            {
                //--------------------------------------RGB疊圖---------------------------------------------------------------------
                Bitmap myimage = new Bitmap(this.CCDImage.Width + 5, this.CCDImage.Height + 10);
                Graphics gg = Graphics.FromImage(myimage);
                gg.CopyFromScreen(new Point(this.Location.X + CCDImage.Location.X + 5, this.Location.Y + CCDImage.Location.Y + 28), new Point(CCDImage.Location.X, CCDImage.Location.Y), new Size(CCDImage.Width + 15, CCDImage.Height + 20));
                IntPtr dc1 = gg.GetHdc();
                gg.ReleaseHdc(dc1);
                myimage.Save(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_full_Image_" + $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg");
                Bitmap After_Overlay = Overlay_Rgb(p1, myimage);
                p1 = After_Overlay;
                if (Laser == 8)
                {
                    p1.Save(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_full_Image_" + $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg");
                    p1.Dispose();
                }

            }
            else
            {
                Bitmap myimage = new Bitmap(this.CCDImage.Width + 5, this.CCDImage.Height + 10);
                Graphics gg = Graphics.FromImage(myimage);
                gg.CopyFromScreen(new Point(this.Location.X + CCDImage.Location.X + 5, this.Location.Y + CCDImage.Location.Y + 28), new Point(CCDImage.Location.X, CCDImage.Location.Y), new Size(CCDImage.Width + 15, CCDImage.Height + 20));
                IntPtr dc1 = gg.GetHdc();
                gg.ReleaseHdc(dc1);
                myimage.Save(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_full_Image_" + $"{name}" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg");

            }
            isReady = true;
        }

        private void Update_Dg(int dg)
        {
            textBox9.Text = dg.ToString();
            if (dg > 255)
            {  }
            else if(dg <32)
            { }
            else
            {
                trackBar1.Value = dg;
            }
        }
        private void Update_Gamma(int gamma)
        {            
            textBox10.Text = gamma.ToString();
            if (gamma > 300)
            { }
            else if (gamma < 100)
            { }
            else
            {
                trackBar2.Value = gamma;
            }
        }
        private void Update_Back(int back)
        {
            textBox11.Text = back.ToString();
            if (back > 3)
            { }
            else if (back < 0)
            { }
            else
            {
                trackBar3.Value = back;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Pause = false;
            iTask = 625;
            onlyWhite = true;
            back_number = 0;
            checkBox8.Checked = true;
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            Set_Gamma(trackBar2.Value);
            label33.Text = trackBar2.Value.ToString();
        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            Set_Back(trackBar3.Value);
            label37.Text = trackBar3.Value.ToString();
        }

        private void timer2_roi_Tick(object sender, EventArgs e)
        {

        }

        private void SC_ID_Click(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            List<double> wavelengh = new List<double>();
            List<string> wavelengh_for_save = new List<string>();
            bool is_wavelenght_save = false;
            bool is_havewave = false;
            wavelengh.Clear();
            if (is_havewave == false)
            {
                for (int i = 0; i < 1280; i++)
                {
                    wavelengh.Add(Poly_Coefs[3] * Math.Pow(i, 3) + Poly_Coefs[2] * Math.Pow(i, 2) + Poly_Coefs[1] * i + Poly_Coefs[0]);
                    wavelengh_for_save.Add(wavelengh[i].ToString());
                }
                is_havewave = true;
            }
            if (is_wavelenght_save == false)
            {
                File.WriteAllLines(@"RawData\Auto_Save\" + this_SC_ID + @"\" + this_SC_ID + "本次校正後對應波長" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + "_.txt", wavelengh_for_save.ToArray());
            }
            System.Windows.Forms.DataVisualization.Charting.Series seriesGray = new System.Windows.Forms.DataVisualization.Charting.Series("灰階", 2000);

            //設定座標大小
            this.chart_with_wavelenght.ChartAreas[0].AxisX.Interval = 100;
            this.chart_with_wavelenght.ChartAreas[0].AxisX.Minimum = Math.Round(wavelengh[0],2);
            this.chart_with_wavelenght.ChartAreas[0].AxisX.Maximum = Math.Round(wavelengh[wavelengh.Count-1],2);
            this.chart_with_wavelenght.ChartAreas[0].AxisY.Minimum = 0;
            this.chart_with_wavelenght.ChartAreas[0].AxisY.Maximum = 300;


            //設定標題
            this.chart_with_wavelenght.Titles.Clear();
            this.chart_with_wavelenght.Titles.Add("S01");
            if (checkBox6.Checked)
            {
                this.chart_with_wavelenght.Titles[0].Text = "原始光譜_"+this_SC_ID+"_Hg_Ar";
            }
            else if (checkBox7.Checked)
            {
                this.chart_with_wavelenght.Titles[0].Text = "原始光譜_" + this_SC_ID + "_Laser";
            }
            else if (checkBox8.Checked)
            {
                this.chart_with_wavelenght.Titles[0].Text = "原始光譜_" + this_SC_ID + "_White";
            }
            this.chart_with_wavelenght.Titles[0].ForeColor = Color.Black;
            this.chart_with_wavelenght.Titles[0].Font = new System.Drawing.Font("標楷體", 16F);
            // this.chart_original.


            //設定顏色
            seriesGray.Color = Color.Blue;
            
            //設定樣式
            seriesGray.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
           
            //迴圈二
            for (int x = 0; x < wavelengh.Count; x++)
            {
                //給入數據畫圖
                seriesGray.Points.AddXY(Math.Round(wavelengh[x],2), RealTime_Original_Intensity[x]);
                this.chart_with_wavelenght.Series.Clear();
                this.chart_with_wavelenght.Series.Add(seriesGray);
               
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            timer2.Start();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            chart_with_wavelenght.SaveImage(@"Screenshot\" + this_SC_ID + @"\" + this_SC_ID + "_wave_with_intensity_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg", ImageFormat.Bmp);
            MessageBox.Show("截圖完成!");
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string html = File.ReadAllText("demo.html");
            webBrowser1.DocumentText = html;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                //isch1 = true; 
                isch2 = true; isch3 = true; isch5 = true; isch6 = true; isch7 = true; isch8 = true;
               // btn_Ch1.BackColor = Color.Green;
                btn_Ch2.BackColor = Color.Green;
                btn_Ch3.BackColor = Color.Green;
                btn_Ch5.BackColor = Color.Green;
                btn_Ch6.BackColor = Color.Green;
                btn_Ch7.BackColor = Color.Green;
                btn_Ch8.BackColor = Color.Green;
                AllinOneMode.Checked = false;
                CH1_CHECK.Checked = false;
                CH6_CHECK.Checked = false;
            }
            else
            {
                isch1 = false; isch2 = false; isch3 = false; isch5 = false; isch6 = false; isch7 = false; isch8 = false;
                AllinOneMode.Checked = true;
                btn_Ch1.BackColor = Color.LightGray;
                btn_Ch2.BackColor = Color.LightGray;
                btn_Ch3.BackColor = Color.LightGray;
                btn_Ch5.BackColor = Color.LightGray;
                btn_Ch6.BackColor = Color.LightGray;
                btn_Ch7.BackColor = Color.LightGray;
                btn_Ch8.BackColor = Color.LightGray;
                CH1_CHECK.Checked = true;
                CH6_CHECK.Checked = true;
            }
        }

        private void AllinOneMode_CheckedChanged(object sender, EventArgs e)
        {
            if (AllinOneMode.Checked)
            {
                checkBox5.Checked = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                checkBox7.Checked = false; checkBox8.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                checkBox6.Checked = false; checkBox8.Checked = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                checkBox6.Checked = false; checkBox7.Checked = false;
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            int k = 0;
            bool end = false;
            bool isGetLeft = false;
            bool isGetRight = false;
            List<double> baseline = new List<double>();
            List<double> Result = new List<double>();
            List<double> new_baseline = new List<double>();
            int index = 0;
            int left_Point = 0;
            int right_Point = 0;
            List<double> wavelengh = new List<double>();
            List<double> part_of_wavelength = new List<double>();
            wavelengh.Clear();
            baseline.Clear();
            new_baseline.Clear();
            Result.Clear();
            //--------------------------------------------------------------------------------------------------------------------------
            for (int i = 0; i < 1280; i++)
            {
                wavelengh.Add(Poly_Coefs[3] * Math.Pow(i, 3) + Poly_Coefs[2] * Math.Pow(i, 2) + Poly_Coefs[1] * i + Poly_Coefs[0]);
            }
            //--------------------------------------------------------------------------------------------------------------------------
            while (end==false)
            {

                if (wavelengh[k] > Convert.ToDouble(textBox5.Text)&&isGetLeft==false)
                {
                    left_Point = k;
                    isGetLeft = true;
                }
                else if (wavelengh[k] > Convert.ToDouble(textBox6.Text)&&isGetRight==false)
                {
                    right_Point = k;
                    isGetRight = true;
                }
                else if (isGetRight && isGetLeft)
                {
                    end = true;
                }
                else
                { k++; }
                
            }
            //--------------------------------------------------------------------------------------------------------------------------
            for (int i = left_Point; i < right_Point; i++)
            {
                baseline.Add(RealTime_Original_Intensity[i]);
                part_of_wavelength.Add(wavelengh[i]);
            }
            //-------------------------------------------------------------------------------------------------------------------------
            Result = Math_Methods.Polynomial_Fitting(part_of_wavelength,baseline,3);
            for (int i = 0; i < part_of_wavelength.Count; i++)
            {
                new_baseline.Add(Result[3] * Math.Pow(part_of_wavelength[i],3) + Result[2] * Math.Pow(part_of_wavelength[i], 2) + Result[1] * part_of_wavelength[i] + Result[0]);
            }
           index = new_baseline.IndexOf(new_baseline.Min());
            label43.Text = Math.Round(part_of_wavelength[index],2).ToString() + " ± Δ" + textBox7.Text + " nm";
            Baseline_wavelength_of_minPoint = part_of_wavelength[index];
            Baseline_of_leftPoint = Baseline_wavelength_of_minPoint - Convert.ToDouble(textBox7.Text);
            Baseline_of_rightPoint = Baseline_wavelength_of_minPoint + Convert.ToDouble(textBox7.Text);

            //---------------------------------------------------區分較靠近哪個以0.5為單位的數值-------------------------------------
            if (Math.Round(Baseline_of_leftPoint) > Baseline_of_leftPoint)
            {
                if (Math.Round(Baseline_of_leftPoint)-0.25 > Baseline_of_leftPoint)
                {
                    Baseline_of_leftPoint = Math.Round(Baseline_of_leftPoint) - 0.5;
                }
                else
                {
                    Baseline_of_leftPoint = Math.Round(Baseline_of_leftPoint);
                }
            }
            else 
            {
                if (Math.Round(Baseline_of_leftPoint) + 0.25 > Baseline_of_leftPoint)
                {
                    Baseline_of_leftPoint = Math.Round(Baseline_of_leftPoint);
                }
                else
                {
                    Baseline_of_leftPoint = Math.Round(Baseline_of_leftPoint) + 0.5;
                }
            }

            if (Math.Round(Baseline_of_rightPoint) > Baseline_of_rightPoint)
            {
                if (Math.Round(Baseline_of_rightPoint) - 0.25 > Baseline_of_rightPoint)
                {
                    Baseline_of_rightPoint = Math.Round(Baseline_of_rightPoint) - 0.5;
                }
                else
                {
                    Baseline_of_rightPoint = Math.Round(Baseline_of_rightPoint);
                }
            }
            else
            {
                if (Math.Round(Baseline_of_rightPoint) + 0.25 > Baseline_of_rightPoint)
                {
                    Baseline_of_rightPoint = Math.Round(Baseline_of_rightPoint);
                }
                else
                {
                    Baseline_of_rightPoint = Math.Round(Baseline_of_rightPoint) + 0.5;
                }
            }
            //-------------------------------------------------------------------------------------------------------
        }

        private void button21_Click(object sender, EventArgs e)
        {
            JsonReCreate create = new JsonReCreate();
            create.ShowDialog();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            List<string> Scale = new List<string>();
            Scale.Add(textBox8.Text);
            File.WriteAllLines(@"LaserScale\" + "AutoScale.txt", Scale);
        }

        #region 視窗截圖
        public void CaptureWindow()
        {

            /* 取得目標視窗的 Handle
             * 需要加上 using System.Diagnostics;
             */
            IntPtr hwnd = FindWindow(null, textBox4.Text);

            /* 取得該視窗的大小與位置 */
            CapRectangle bounds = new CapRectangle();

            //GetWindowRect(process[0].MainWindowHandle, ref bounds);
            GetWindowRect(hwnd, ref bounds);
            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;

            SetWindowPos(hwnd, -1, 0, 0, 0, 0, 1 | 2);//截圖視窗至於最上層

            /* 抓取截圖 */
            Bitmap screenshot = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            Graphics gfx = Graphics.FromImage(screenshot);
            gfx.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);

            /* 利用 PictureBox 顯示出來 */
            Directory.CreateDirectory(@"Screenshot\"+this_SC_ID +@"\"+ @"Result\");
            string path = @"Screenshot\" + this_SC_ID + @"\" + @"Result\" +"Result_"+ DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".jpg";
            //@"Screenshot\" + DateTime.Now.ToString("yyyyMMdd HHmmss") + ".jpg";
            Image save = (Image)screenshot;
            save.Save(path);
            MessageBox.Show("存檔完成");
        }
        #endregion
        private void btn_Ch1_Click_1(object sender, EventArgs e)
        {
            isch1 = true;
            btn_Ch1.BackColor = Color.Green;
        }
        #region 控件隨窗體大小調整
        

        // <summary>

        /// 自适应窗体大小的方法，递归

        /// </summary>

        /// <param name="theControl"></param>

        private void AutoResize(Control theControl)  //核心代码
        {

            //------------------ 控制控件自适应UI大小---------------------------------

            foreach (Control OneCon in theControl.Controls)   //根据原大小等比例放大

            {

                OneCon.Left = Convert.ToInt32((OneCon.Left * fRatio_Width));

                OneCon.Top = Convert.ToInt32((OneCon.Top * fRatio_Height));

                OneCon.Width = Convert.ToInt32(OneCon.Width * fRatio_Width);

                OneCon.Height = Convert.ToInt32(OneCon.Height * fRatio_Height);

                if (OneCon.Controls.Count > 0)

                {

                    AutoResize(OneCon);//如Form内还有控件中的控件，此处则可递归改变大小，此处为重点中的重点  //2019-3-21 10:50:42更新

                }

                // 自动更改字体的大小  //2020-03-23 19:21:50 更新

                // 当前控件（OneCon.Font.Size） 乘以(8)  长宽比例中最小值 (Math.Min(fRatio_Width, fRatio_Height)

                //Math.Round(数值, 2) //保留两位有效数字

                float currentSize = float.Parse(Math.Round(OneCon.Font.Size * (Math.Min(fRatio_Width, fRatio_Height)), 2).ToString());

                if (currentSize != FontSizeMin)
                {
                    OneCon.Font = new Font(OneCon.Font.Name, currentSize, OneCon.Font.Style, OneCon.Font.Unit);
                }

                /* else

                     OneCon.Font = new Font(OneCon.Font.Name, FontSizeMin, OneCon.Font.Style, OneCon.Font.Unit);*/

                //某种程度上来说 if 语句都可以改成三目运算符， 举例： m=a>b?a:b; 表示先判断a是否大于b，若a>b,则将a的值赋给m，若不符合a>b,则将b的值赋给m

                //float NowSize = currentSize > FontSizeMin ? currentSize : FontSizeMin;

                //OneCon.Font = new Font(OneCon.Font.Name, NowSize, OneCon.Font.Style, OneCon.Font.Unit);

                

            }
            this.chart_original.ChartAreas["ChartArea1"].AxisX.Maximum = 1280;


        }
        private void Form1_ResizeEnd(object sender, EventArgs e)
        {
            //窗体改变后Form的宽度与高度及宽比例和高比例

            fRatio_Width = (this.Size.Width / (float)iForm_ResizeBefore_Width);

            fRatio_Height = (this.Size.Height / (float)iForm_ResizeBefore_Height);

            AutoResize(this); //2019-3-21 10:50:42更新



            ////获得下一次窗体改变前的Form的宽度与高度

            iForm_ResizeBefore_Width = this.Size.Width;

            iForm_ResizeBefore_Height = this.Size.Height;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            //窗体改变后Form的宽度与高度及宽比例和高比例

            fRatio_Width = (this.Size.Width / (float)iForm_ResizeBefore_Width);

            fRatio_Height = (this.Size.Height / (float)iForm_ResizeBefore_Height);

            AutoResize(this); //2019-3-21 10:50:42更新



            ////获得下一次窗体改变前的Form的宽度与高度

            iForm_ResizeBefore_Width = this.Size.Width;

            iForm_ResizeBefore_Height = this.Size.Height;
        }

        #endregion

        private void button10_Click(object sender, EventArgs e)
        {
            CaptureWindow();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }
      
    }
}


