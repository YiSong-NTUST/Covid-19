using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SpertroApp
{
    class Step_1
    {
      // public int progressBar = 0;


        public static IDictionary<string, int> RoiScan(Bitmap input_image0)
        {
            // Scan_Line.Visible = true;
            Bitmap input_image = new Bitmap(input_image0.Width, input_image0.Height);
            input_image = input_image0;
            IDictionary<string, int> ROI = new Dictionary<string, int>();
            /*
            ROI.Add("x", 0);//不變
            ROI.Add("y", 0);
            ROI.Add("w", wid);
            ROI.Add("h", 0);//不變*/
            int progressBar = 0;
            int now_y = 0;//總共有 input_image.Height-20個
            /*
             * 原本->從0掃描到 input_image.Height-20
             * 現在->  從clip掃描到 input_image.Height-20-clip           
             */

            int roi_fixHeight = 20;//之後可設為外部設定
            int Pixel_x = 0;//正在被掃描的點
            int Pixel_y = 0;
            int sum_gray = 0;
            int clip = 380; //被剪去的部分,進而加速,之後可設為外部設定380
            // int [] sum4eachROI=new int[];
            List<int> sum4eachROI = new List<int>();

            now_y = clip;

            while (now_y < (input_image0.Height - roi_fixHeight- clip))
            {
                for (Pixel_x = 0; Pixel_x < input_image0.Width; Pixel_x++)//一直是0->input_image.Width
                {
                    for (Pixel_y = now_y; Pixel_y < (now_y + roi_fixHeight); Pixel_y++)//初始值為now_y掃到now_y+20
                    {
                        Color p0 = input_image.GetPixel(Pixel_x, Pixel_y);
                        int R = p0.R, G = p0.G, B = p0.B;
                        int gray = (R * 313524 + G * 615514 + B * 119538) >> 20;

                        sum_gray = sum_gray + gray;
                    }

                }
                // Add parts to the list.
                sum4eachROI.Add(sum_gray);
                //sum4eachROI[now_y] = sum_gray;

                sum_gray = 0;
                now_y++;
                progressBar++;
          //      Form f1 = new Form1();

           /*     f1.progressROI.
                progressBar1.Value += progressBar1.Step;//讓進度條增加一次*/

            }
            //最大值的INDEX?
            // IEnumerable<int> MAX_y = sum4eachROI.OrderByDescending(index => index).Take(1);
            //IEnumerable<int> MAX_y = sum4eachROI.Select((m, index) => new { index, m }).OrderByDescending(n => n.m).Take(1);
            var MAX_y = sum4eachROI.IndexOf(sum4eachROI.Max())+clip;

            ROI.Add("x", 0);//不變
            ROI.Add("y", Convert.ToInt32(MAX_y));
            ROI.Add("w", input_image0.Width);
            ROI.Add("h", roi_fixHeight);//不變
         //   this.Invoke(formcontrl, DrawCanvas.Top, 0, Context);

            Console.WriteLine(progressBar);
            return ROI;
        }





    }
}
