using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//視窗截圖
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Threading;

namespace SpertroApp
{
    public partial class Report : Form
    {
        private Form1 f1;
        bool isReady = false;
        public Report(Form1 form)
        {
            InitializeComponent();
            f1 = form;
        }

        private void Report_Load(object sender, EventArgs e)
        {
            textBox2.Text = f1.this_SC_ID;
            
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

        #region 視窗截圖
        public void CaptureWindow()
        {

            /* 取得目標視窗的 Handle
             * 需要加上 using System.Diagnostics;
             */
            IntPtr hwnd = FindWindow(null, "Report");

            /* 取得該視窗的大小與位置 */
            CapRectangle bounds = new CapRectangle();

            //GetWindowRect(process[0].MainWindowHandle, ref bounds);
            GetWindowRect(hwnd, ref bounds);
            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;

            SetWindowPos(hwnd, -1, 0, 0, 0, 0, 1 | 2);//截圖視窗至於最上層
            
            if (isReady)
            {
                /* 抓取截圖 */
                Bitmap screenshot = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                Graphics gfx = Graphics.FromImage(screenshot);
                gfx.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                Directory.CreateDirectory(@"Screenshot\Report\" + f1.this_SC_ID);
                string path = @"Screenshot\Report\" + f1.this_SC_ID + @"\" + "Report_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm") + ".jpg";
                Image save = (Image)screenshot;
                save.Save(path);
                MessageBox.Show("截圖完成");
            }
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox2.Text;
            StreamWriter sw = new StreamWriter(@"Report\"+ name +"_"+ "Report"+"_"+DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".txt");
            //Write a line of text
            sw.WriteLine(textBox1.Text);
            //Close the file
            sw.Close();
            MessageBox.Show("存檔成功");
        }

        private void Report_Shown(object sender, EventArgs e)
        {
            /*
            string report = "λ , ΔλL(nm)" + "\t" + "Pixel" + "\t" + "ΔP" + "\t" + "Δxm" + "\t" + "Δλm" + "\t" + "Δλsc" + "\t" + "ΔPsc" + "\t" + "ΔXsc" + "  " + "Δλrms" + "  " + "  " + "Δλstd" + "  " + "  " + "Δλstd%" + "  " + "  " + "ΔXrms" + "  " + "  " + "ΔXstd" + "  " + "  " + "ΔXstd%" + "\r\n";
            if (f1.isHg_Ar)
            {
                report += "\r\n" + "-----------------------------------------------------------------------------------------HG-----------------------------------------------------------------------------------------------\n";
                for (int i = 0; i < f1.Wavelength.Count; i++)
                {                   
                    report += "\r\n" + f1.Wavelength[i].ToString("f2") + " , " + f1.deltaL[i].ToString("f2") //λ , ΔλL(nm)
                        + "\t" + f1.pixel_M[i].ToString("f2") //Pixel
                        + "\t" + f1.pFWHM[i].ToString("f2") //ΔP
                        + "\t" + (f1.pFWHM[i] * 3.75).ToString("f2") //Δx
                        + "\t" + f1.wFWHM[i].ToString("f2") //Δλm
                        + "\t" + f1.LambdaSC[i].ToString("f2")//Δλsc
                        + "\t" + f1.DeltaPsc[i].ToString("f2") //ΔPsc
                        + "\t" + (f1.DeltaPsc[i] * 3.75).ToString("f2")//ΔXsc
                        + "\t" + f1.LambdaSC_RMS.ToString("f2")//ΔλRMS
                        + "\t" + "  " + f1.LambdaSC_STD.ToString("f2")//Δλstd
                        + "\t" + "  " + f1.LambdaSC_STD_100.ToString("f2")//ΔλRMS%
                        + "\t" + "  " + "  " + f1.LambdaSC_Spot_RMS.ToString("f2")//ΔXRMS
                        + "\t" + "  " + "  " + f1.LambdaSC_Spot_STD.ToString("f2")//ΔXstd
                        + "\t" + "  " + "  " + f1.LambdaSC_Spot_STD_100.ToString("f2");//ΔXRMS%

                }
            }
            else 
            {
                report += "\r\n" + "-----------------------------------------------------------------------------------------Laser-----------------------------------------------------------------------------------------------\n";
                for (int i = 0; i < f1.Wavelength.Count; i++)
                {
                    report += "\r\n" + f1.Wavelength[i].ToString("f2") + " , " + f1.deltaL[i].ToString("f2") //λ , ΔλL(nm)
                        + "\t" + f1.pixel_M[i].ToString("f2") //Pixel
                        + "\t" + f1.pFWHM[i].ToString("f2") //ΔP
                        + "\t" + (f1.pFWHM[i] * 3.75).ToString("f2") //Δx
                        + "\t" + f1.wFWHM[i].ToString("f2") //Δλm
                        + "\t" + f1.LambdaSC[i].ToString("f2")//Δλsc
                        + "\t" + f1.DeltaPsc[i].ToString("f2") //ΔPsc
                        + "\t" + (f1.DeltaPsc[i] * 3.75).ToString("f2")//ΔXsc
                        + "\t" + f1.LambdaSC_RMS.ToString("f2")//ΔλRMS
                        + "\t" + "  " + f1.LambdaSC_STD.ToString("f2")//Δλstd
                        + "\t" + "  " + f1.LambdaSC_STD_100.ToString("f2")//ΔλRMS%
                        + "\t" + "  " + "  " + f1.LambdaSC_Spot_RMS.ToString("f2")//ΔXRMS
                        + "\t" + "  " + "  " + f1.LambdaSC_Spot_STD.ToString("f2")//ΔXstd
                        + "\t" + "  " + "  " + f1.LambdaSC_Spot_STD_100.ToString("f2");//ΔXRMS%

                }
            }
            report += "\r\n" + "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"
                            + "\r\n" + "a0 = " + f1.Poly_Coefs[0]
                            + "\r\n" + "a1 = " + f1.Poly_Coefs[1]
                            + "\r\n" + "a2 = " + f1.Poly_Coefs[2]
                            + "\r\n" + "a3 = " + f1.Poly_Coefs[3]
                            + "\n";
            */
            if (f1.Laser_And_Hg_Ar_CHK.Checked)
            {
                textBox1.AppendText(f1.report_Memory);
            }
            else
            {
                textBox1.AppendText(f1.report);
            }
            isReady = true;
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CaptureWindow();
        }
    }
}
