using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GitHub.secile.Video;

namespace SpertroApp
{
    public partial class Form_camp2 : Form
    {
        static Size Read_Size = new Size(1280, 960);
        UsbCamera camera = new UsbCamera(0, Read_Size);




        public Form_camp2()
        {
            InitializeComponent();
        }

        public Form_camp2(int Gamma, int Back)
        {
            InitializeComponent();
            trackBar1.Value = Gamma;
            trackBar2.Value = Back;

        }
        private void Form_camp2_Load(object sender, EventArgs e)
        {
            int cameraIndex = 0;
            // check format.
            string[] devices = UsbCamera.FindDevices();
            if (devices.Length == 0) return; // no camera.

            GitHub.secile.Video.UsbCamera.VideoFormat[] formats = UsbCamera.GetVideoFormat(cameraIndex);
            //for (int i = 0; i < formats.Length; i++) Console.WriteLine("{0}:{1}", i, formats[i]);

            // create usb camera and start.
            camera = new UsbCamera(cameraIndex, formats[11]); //1280*960
            Console.WriteLine(formats[8] + "\n" + formats[9] + "\n" + formats[10] + "\n" + formats[11]);
            timer1.Start();
            camera.Start();


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
            set_camera_prop("gain", 32);

            //讀取
           /*trackBar1.Value= read_camera_prop("gamma");
            trackBar2.Value = read_camera_prop("back");
            trackBar1.Value = FI
            trackBar2.Value = read_camera_prop("back");*/
        }


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

                prop.SetValue(DirectShow.CameraControlFlags.Manual, prop.Default);

            }


        }

        int read_camera_prop(string item)
        {
            UsbCamera.PropertyItems.Property prop;
            switch (item)
            {
                case "exp":
                    prop = camera.Properties[DirectShow.CameraControlProperty.Exposure];

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



            

            return prop.Default;
        }



        private void timer1_Tick(object sender, EventArgs e)
        {
            Bitmap myBitmap = camera.GetBitmap();
            myBitmap.RotateFlip(RotateFlipType.RotateNoneFlipX);

            //    scaleX = camera.Size.Width / CCDImage.Width;
            //  scaleY = camera.Size.Height / CCDImage.Height;



            System.Drawing.Imaging.PixelFormat format = myBitmap.PixelFormat;


            CCDImage.Image = myBitmap;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            camera.Stop();
            Form_camp FC = new Form_camp(trackBar1.Value, trackBar2.Value);
            FC.Show();

            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            camera.Stop();
            Form1 F1 = new Form1(trackBar1.Value, trackBar2.Value);
            F1.Show();

            this.Close();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            set_camera_prop("gamma", trackBar1.Value);
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            set_camera_prop("back", trackBar2.Value);
        }

        private void Form_camp2_FormClosing(object sender, FormClosingEventArgs e)
        {
           /* Form_camp FP2 = new Form_camp(trackBar1.Value, trackBar2.Value);
            FP2.Show();*/
        }
    }
}
