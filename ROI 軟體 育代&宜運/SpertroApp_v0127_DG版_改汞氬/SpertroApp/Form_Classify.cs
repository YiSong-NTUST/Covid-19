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
    public partial class Form_Classify : Form
    {
        static Size Read_Size = new Size(1280, 960);
        UsbCamera camera = new UsbCamera(0, Read_Size);
        public Form_Classify()
        {
            InitializeComponent();
        }

        private void Form_Classify_Load(object sender, EventArgs e)
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
            timer1.Start();
            camera.Start();
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
    }
}
