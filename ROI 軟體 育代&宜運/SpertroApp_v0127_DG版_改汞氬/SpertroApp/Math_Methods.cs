using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MathNet.Numerics.LinearAlgebra; //<----之後正式的高斯就不要這個
using MathWorks.MATLAB.NET.Arrays;
using sg;
using PolyFit;
using gaussfit;
using SpertroApp;
using GitHub.secile.Video;
using System.Threading;
using autofindpeaks;
using Plot;
using PolyFit_NPlot;
using Poly_And_Plot;



namespace SpectroChipApp
{

    class Math_Methods
    {


        //   Original_Intensity ---去基線-----> Pure_Intensity ---SG--->SG_Intensity ---高斯擬和--->Gaus_Intensity

        //@(已完成)Control_EXP(Bitmap)回傳和目標差多遠的狀態
        //@(已完成)Pure_Intensity = Remove_BaseLine  (List Original_Intensity, List Dark_Intensity)
        //@(已完成)SG_Intensity= SG_Fitting(List<double> Original_Intensity, int order, int window)
        //@(已完成)Auto_Gaussian(Intensity)
        /*
         *private Find_Gaussia n_Region()
         * private Gaussian_Clip(intensity, xFront,xBack,out intensity_clip, out pixel_clip)
         * Gaussian(IntensitySG, int fitDatasCount, int order, out double[] Gausian_y, out double xMax, out double yMax, out 標準差 )
       */

        //@(已完成)coefs = Polynomial (w, p) 
        //@(已完成)LamdaFWHM=getLamdaFWHM(pFWHM, coefs) <-微分帶進去
        //@(已完成)waveLength=WL_Calibrate(Pixel, parameter)
        //@double=getRMSE(list1,List2)


        /*-------------------------------------------------------Public-------------------------------------------------------*/

        public static double getRMSE(List<double> Original, List<double> Fit)
        {
            double RMSE = 0;
            double SUM = 0;
            List<double> Dev = Remove_BaseLine(Fit, Original);


            for (int i = 0; i < Dev.Count; i++)
            {
                Dev[i] *= Dev[i];
                SUM = Dev[i] + SUM;
            }

            RMSE = Math.Sqrt(SUM / Dev.Count);
            return RMSE;
        }



        public static List<double> get_Original_Intensity_byte(byte[] input_image_buffer, int W, int H)
        {
            //this.chart2.Series.Clear();

            //    W = input_image.Width;
            //  H = input_image.Height;
            //Bitmap image_roi_for_gray = new Bitmap(w, h);




            int Pixel_x = 0;//正在被掃描的點
            int Pixel_y = 0;
            double[] ARed = new double[W];
            double[] AGreen = new double[W];
            double[] ABlue = new double[W];
            double[] AGray = new double[W];
            double[] IntensityRed = new double[W];
            double[] IntensityGreen = new double[W];
            double[] IntensityBlue = new double[W];
            double[] IntensityGray = new double[W];

            //var sgf = MathNet.Filtering.OnlineFilter.CreateDenoise(window);


            //    System.Windows.Forms.DataVisualization.Charting.Series seriesGray = new System.Windows.Forms.DataVisualization.Charting.Series("灰階", 2000);

            // System.Windows.Forms.DataVisualization.Charting.Series seriesSG1 = new System.Windows.Forms.DataVisualization.Charting.Series("SG", 2000);
            int gray = 0;

            for (Pixel_x = 0; Pixel_x < W; Pixel_x++)
            {
                for (Pixel_y = 0; Pixel_y < H; Pixel_y++)
                {
                    gray = input_image_buffer[Pixel_y * W + Pixel_x] + gray;
                    //   AGray[Pixel_x] = AGray[Pixel_x] + gray;

                }
                IntensityGray[Pixel_x] = gray / H;
                gray = 0;
            }

            return IntensityGray.ToList();
        }


        public static List<double> get_Original_Intensity(Bitmap input_image)
        {
            //this.chart2.Series.Clear();
            int W;
            int H;
            W = input_image.Width;
            H = input_image.Height;
            //Bitmap image_roi_for_gray = new Bitmap(w, h);



            // Bitmap im1 = new Bitmap(w, h);//讀出原圖X軸 pixel
            Bitmap im1 = new Bitmap(W, H);//讀出原圖X軸 pixel
            int Pixel_x = 0;//正在被掃描的點
            int Pixel_y = 0;
            double[] ARed = new double[W];
            double[] AGreen = new double[W];
            double[] ABlue = new double[W];
            double[] AGray = new double[W];
            double[] IntensityRed = new double[W];
            double[] IntensityGreen = new double[W];
            double[] IntensityBlue = new double[W];
            double[] IntensityGray = new double[W];

            //var sgf = MathNet.Filtering.OnlineFilter.CreateDenoise(window);


            //    System.Windows.Forms.DataVisualization.Charting.Series seriesGray = new System.Windows.Forms.DataVisualization.Charting.Series("灰階", 2000);

            // System.Windows.Forms.DataVisualization.Charting.Series seriesSG1 = new System.Windows.Forms.DataVisualization.Charting.Series("SG", 2000);

            for (Pixel_x = 0; Pixel_x < W; Pixel_x++)
            {
                for (Pixel_y = 0; Pixel_y < H; Pixel_y++)
                {
                    //int xx = Pixel_x + 1;
                    //int yy = Pixel_y + 1;

                    //先把圖變灰階

                    Color p0 = input_image.GetPixel(Pixel_x, Pixel_y);//太快會閃退，全世界都在用image_roi
                    int R = p0.R, G = p0.G, B = p0.B;
                    int gray = (R * 313524 + G * 615514 + B * 119538) >> 20;
                    Color p1 = Color.FromArgb(gray, gray, gray);
                    //im1.SetPixel(i, j, p1);
                    ARed[Pixel_x] = ARed[Pixel_x] + R;
                    AGreen[Pixel_x] = AGreen[Pixel_x] + G;
                    ABlue[Pixel_x] = ABlue[Pixel_x] + B;
                    AGray[Pixel_x] = AGray[Pixel_x] + gray;
                }
                IntensityRed[Pixel_x] = ARed[Pixel_x] / H;//平均
                IntensityGreen[Pixel_x] = AGreen[Pixel_x] / H;//平均
                IntensityBlue[Pixel_x] = ABlue[Pixel_x] / H;//平均
                IntensityGray[Pixel_x] = AGray[Pixel_x] / H;//平均
            }

            return IntensityGray.ToList();
        }


        public static List<double> getLamdaFWHM(List<double> PixelFWHM, List<double> p, List<double> pixel_M)
        {
            List<double> LamdaFWHM = new List<double>();
            for (int i = 0; i < pixel_M.Count; i++)
            {
                // LamdaFWHM.Add((4 * p[4] * Math.Pow(pixel_M[i], 3) + 3 * p[3] * Math.Pow(pixel_M[i], 2) + 2 * p[2] * pixel_M[i] + p[1])* PixelFWHM[i]);
                LamdaFWHM.Add((3 * p[3] * Math.Pow(pixel_M[i], 2) + 2 * p[2] * pixel_M[i] + p[1]) * PixelFWHM[i]);
            }
            return LamdaFWHM;
        }

        public static List<double> getDeltaPsc(List<double> DeltaSc, List<double> p, List<double> pixel_M)
        {
            List<double> DeltaPsc = new List<double>();
            for (int i = 0; i < pixel_M.Count; i++)
            {
                DeltaPsc.Add(DeltaSc[i] / (3 * p[3] * Math.Pow(pixel_M[i], 2) + 2 * p[2] * pixel_M[i] + p[1]));
            }
            return DeltaPsc;
        }

        public static List<double> WL_Calibrate(List<double> X_Pixels, List<double> p)
        {
            List<double> X_Wave = new List<double>();

            for (int i = 0; i < X_Pixels.Count; i++)
                X_Wave[i] = p[4] * Math.Pow(X_Wave[i], 4) + p[3] * Math.Pow(X_Wave[i], 3) + p[2] * Math.Pow(X_Wave[i], 2) + p[1] * X_Wave[i] + p[0];


            return X_Wave;
        }


        public static List<double> Polynomial_Fitting(List<double> X_Pixles, List<double> Y_Wavelength, int order)
        {

            MWArray order_M = (MWNumericArray)order;
            MWArray X_Pixles_M = (MWNumericArray)X_Pixles.ToArray();
            MWArray Y_Wavelength_M = (MWNumericArray)Y_Wavelength.ToArray();



            List<double> Coef_List = new List<double>();
            PF_NP pf_np = new PF_NP();

            MWArray Fit = pf_np.PolyFit_NPlot(X_Pixles_M, Y_Wavelength_M, order_M);


            return MWArray2Array(Fit).ToList();

        }

        private static List<double> MWArray2List(MWArray Array_M)
        {
            List<double> List = new List<double>();

            double[,] dd;
            dd = (double[,])((MWNumericArray)Array_M).ToArray();
            double[] d = new double[dd.Length];
            for (int i = 0; i < dd.Length; i++) d[i] = dd[0, i];
            List = d.ToList();// ((double[,])((MWNumericArray)yout).ToArray()).ToList<double>();           
            return List;
        }

        private static double[] MWArray2Array(MWArray Array_M)
        {


            double[,] dd;
            dd = (double[,])((MWNumericArray)Array_M).ToArray();
            double[] d = new double[5];//預設大小是5，用來放POLY擬和後的參數
            for (int i = 0; i < dd.Length; i++) d[i] = dd[0, dd.Length - i - 1];

            return d;
        }

        private static double[] MWArray2Array2(MWArray Array_M, int n)
        {
            double[,] dd;
            dd = (double[,])((MWNumericArray)Array_M).ToArray();
            double[] d = new double[n];//預設大小是5，用來放POLY擬和後的參數
            for (int i = 0; i < dd.Length; i++)
            {
                d[i] = dd[0, i];
            }

            return d;
        }




        public static List<int> Control_EXP(byte[] image_buffer)
        {
            int speed = 0;
            Bitmap bmp = Form1.BufferToImage(image_buffer);

            int Max = Bitmap2List(bmp).Max();



            if (Max > 220)
            {
                if (Max > 240) /*超亮*/
                {
                    speed = 2;//亮度剖半
                }
                else/*一點太亮191~220*/
                {
                    speed = 1;//微調一格
                }

            }
            else if (Max < 190)
            {
                if (Max < 50)  /*根本看不到0~15*/
                {

                    speed = -3;

                }
                else if (Max < 100)/*超暗*/
                {

                    speed = -2;

                }
                else/*一點太暗100~170*/
                {
                    speed = -1; //微調一格
                }


            }
            else
            {
                speed = 0;
            }



            List<int> result = new List<int>();
            result.Add(speed);//[0] speed
            result.Add(Max);//[1] 此時的值

            return result;
        }



        /*     public static List<int> Control_EXP_Max_Index(byte[] image_buffer) //step =0 1016
             {
                 int over_expose = 0;
                 Bitmap bmp = Form1.BufferToImage(image_buffer);


                 int Max_Intensity = Bitmap2List(bmp).Max();




                 if (Max_Intensity > 225)
                 {
                     over_expose = 1;

                 }

                 int Index_of_Max = Bitmap2List(bmp).IndexOf(Max_Intensity);

                 List<int> result = new List<int>();
                 result.Add(over_expose);//[0] 是否過曝  0:
                 result.Add(Max_Intensity);//[0] 此時的值
                 result.Add(Index_of_Max);//[1] 發生的位置

                 return result;
             }*/

        /*    public static List<int> Get_Max_from_(byte[] image_buffer) //step =1  1016
            {
                Bitmap bmp = Form1.BufferToImage(image_buffer);


                int Max_Intensity = Bitmap2List(bmp).Max();

                int Index_of_Max = Bitmap2List(bmp).IndexOf(Max_Intensity);

                List<int> result = new List<int>();

                result.Add(Max_Intensity);//[0] 此時的值
                result.Add(Index_of_Max);//[1] 發生的位置

                return result;
            }*/


        public static List<double> Control_EXP_List(List<double> Input_List)
        {
            int speed = 0;


            double Max = Input_List.Max();



            if (Max > 220)
            {
                if (Max > 240) /*超亮*/
                {
                    speed = 2;//亮度剖半
                }
                else/*一點太亮191~220*/
                {
                    speed = 1;//微調一格
                }

            }
            else if (Max < 190)
            {
                if (Max < 50)  /*根本看不到0~50*/
                {

                    speed = -3;

                }
                else if (Max < 180)/*超暗50~180*/
                {

                    speed = -2;

                }
                else/*一點太暗180~190*/
                {
                    speed = -1; //微調一格
                }


            }
            else
            {
                speed = 0;
            }



            List<double> result = new List<double>();
            result.Add(speed);//[0] speed
            result.Add(Max);//[1] 此時的值

            return result;
        }


        public static List<double> Remove_BaseLine(List<double> Original_Intensity, List<double> Dark_Intensity)
        {
            int Data_Length = Original_Intensity.Count;
            List<double> Pure_Intensity = new List<double>(Data_Length);

            for (int i = 0; i < Data_Length; i++) Pure_Intensity.Add(Original_Intensity[i] - Dark_Intensity[i]);

            return Pure_Intensity;
        }
        public static List<double> List_Add(List<double> Original_Intensity, List<double> New_Intensity)
        {
            int Data_Length = Original_Intensity.Count;
            List<double> Sum_Intensity = new List<double>(Data_Length);

            for (int i = 0; i < Data_Length; i++) Sum_Intensity.Add(Original_Intensity[i] + New_Intensity[i]);

            return Sum_Intensity;
        }

        public static List<double> List_Div(List<double> Original_Intensity, int divisor)
        {
            int Data_Length = Original_Intensity.Count;
            List<double> Div_Intensity = new List<double>(Data_Length);

            for (int i = 0; i < Data_Length; i++) Div_Intensity.Add(Original_Intensity[i] / divisor);

            return Div_Intensity;
        }


        public static Dictionary<string, List<double>> Auto_Gaussian(List<double> SG_Intensity)
        {

            Dictionary<string, List<double>> Gaus_Pixel_Intensity_Parameter_Set = new Dictionary<string, List<double>>();

            /*找要fitting的區間*/
            int xFront;
            int xBack;
            Find_Gaussian_Region(SG_Intensity, out xFront, out xBack);

            if (xBack >= SG_Intensity.Count) xBack = SG_Intensity.Count - 1;

            /*切下要fitting的區間*/
            List<double> SG_Intensity_clip; /**/
            List<double> Pixel_clip;
            Gaussian_Clip(SG_Intensity, xFront, xBack, out SG_Intensity_clip, out Pixel_clip);

            /*對該區間做高斯擬和*/
            List<double> Gaus_Intensity_Clip;
            double Pixel_Max;
            double Intensity_Max;
            double SD;

            //    Gaussian(SG_Intensity_clip, SG_Intensity_clip.Count, 3, out Gaus_Intensity_Clip, out Pixel_Max, out Intensity_Max, out SD);
            Gaussian_M(Pixel_clip, SG_Intensity_clip, 0.6, out Gaus_Intensity_Clip, out Pixel_Max, out Intensity_Max, out SD);
            /*計算半高全寬ΔP fwhm*/
            double pFWHM;
            pFWHM = 2 * Math.Sqrt(2 * Math.Log(2)) * SD;

            /*把參數放進參數表*/
            List<double> Parameters_List = new List<double>();
            Parameters_List.Add(Pixel_Max);
            Parameters_List.Add(Intensity_Max);
            Parameters_List.Add(SD);
            Parameters_List.Add(pFWHM);
            Parameters_List.Add(xFront);
            Parameters_List.Add(xBack);


            /*把結果放進字典*/
            Gaus_Pixel_Intensity_Parameter_Set.Add("Pixel", Gaus_Intensity_Clip);
            Gaus_Pixel_Intensity_Parameter_Set.Add("Intensity", Pixel_clip);
            Gaus_Pixel_Intensity_Parameter_Set.Add("Parameters", Parameters_List);


            /*回傳字典*/
            return Gaus_Pixel_Intensity_Parameter_Set;
            /*"Pixel":Pixel_clip   ->畫圖用
            * "Intensity":Gaus_Intensity_Clip  
            * "Parameters"參數表:  ["Parameters"][0]Pixel_Max     ->畫最大點用**
            *                      ["Parameters"][1] Intensity_Max  ->畫最大點用
            *                      ["Parameters"][2]SD(標準差) 
            *                                         ["Parameters"][3]pFWHM(半高全寬) ->**
            *                                         ["Parameters"][4]xFront  ->畫圖用
            *                                         ["Parameters"][5]xBack ->畫圖用
            */
        }





        public static List<double> SG_Fitting(List<double> Original_Intensity, int order, int window)
        {
            MWArray order_M = (MWNumericArray)order;
            MWArray window_M = (MWNumericArray)window;
            MWArray yin_M = (MWNumericArray)Original_Intensity.ToArray();

            SG sgf = new SG();
            /*
            MWArray yout = sgf.sg(yin_M, order_M, window_M);
            
            double[,] dd;

            dd = (double[,])((MWNumericArray)yout).ToArray();

            double[] d = new double[dd.Length];

            for (int i = 0; i < dd.Length; i++) d[i] = dd[0, i];


            List<double> SG_Intensity = new List<double>();

            SG_Intensity = d.ToList();// ((double[,])((MWNumericArray)yout).ToArray()).ToList<double>();
            */

            List<double> SG_Intensity = new List<double>();
            return SG_Intensity;


            //List<double> IntensitySG1 = Math_Methods.SG_Fitting(IntensityGray.ToList<double>(),3,41);

        }

        /*-------------------------------------------------------Private-------------------------------------------------------*/
        //Color
        public static List<int> Bitmap2List(Bitmap image)
        {
            List<int> list = new List<int>();

            for (int x = 0; x < image.Width; x++)
                for (int y = 0; y < image.Height; y++)
                    list.Add(Color_RGB2GrayScale(image.GetPixel(x, y)));

            return list;
        }

        private static Color Color_RGB2Gray(Color RGB)
        {
            int R = RGB.R, G = RGB.G, B = RGB.B;
            int gray = (R * 313524 + G * 615514 + B * 119538) >> 20;
            Color GRAY = Color.FromArgb(gray, gray, gray);
            return GRAY;
        }
        private static int Color_RGB2GrayScale(Color RGB)
        {
            int R = RGB.R, G = RGB.G, B = RGB.B;
            int gray = (R * 313524 + G * 615514 + B * 119538) >> 20;
            return gray;
        }

        //Gaussiane
        private static void Find_Gaussian_Region(List<double> SG_Intensity, out int xFront, out int xBack)
        {
            int xMax;
            double yMax;
            int xHalf;
            double yHalf_theory; //yMax/2
            double yHalf; //最接近yMax/2
            double Half_FWHM_Esitmate;
            int expand = 8;  //擴大選取<-----------找到最適合的區間(找到xFront要小於10之類)


            //output
            xFront = 0;
            xBack = 1280;


            //找最大值
            yMax = SG_Intensity.Max();
            xMax = SG_Intensity.IndexOf(yMax);
            //找y/2
            yHalf_theory = yMax / 2;
            //掃描第一個靠近y/2的點
            yHalf = SG_Intensity.ToArray().OrderBy(y => Math.Abs(y - yHalf_theory)).First();
            xHalf = SG_Intensity.IndexOf(yHalf);
            //求估計的半高全寬的一半
            Half_FWHM_Esitmate = Math.Abs(xMax - xHalf);//if <=0 有問題 

            //求區間
            xFront = xMax - Convert.ToInt32(Math.Round(Half_FWHM_Esitmate)) - expand;
            xBack = xMax + Convert.ToInt32(Math.Round(Half_FWHM_Esitmate)) + expand;
        }
        private static void Gaussian_Clip(List<double> SG_Intensity, int xFront, int xBack, out List<double> SG_Intensity_clip, out List<double> Pixel_clip)
        {
            // SG_Intensity_clip = new List<double>();
            Pixel_clip = new List<double>();

            SG_Intensity_clip = new List<double>(1280);


            if (xFront < 0) xFront = 0;
            SG_Intensity_clip = SG_Intensity.GetRange(xFront, xBack - xFront);
            //SG_Intensity_clip = SG_Intensity.GetRange(5, 10);

            for (int i = 0; i < SG_Intensity_clip.Count; i++)
            {
                Pixel_clip.Add(i + xFront);
            }
            //  throw new NotImplementedException();
        }

        private static void Gaussian(List<double> IntensitySG_List, int fitDatasCount, int order, out List<double> Gausian_y, out double xMax, out double yMax, out double c)
        {

            double[] IntensitySG = IntensitySG_List.ToArray();

            double[,] a = new double[fitDatasCount, order];
            double[] b = new double[fitDatasCount];
            double[] X = new double[order];

            Gausian_y = new List<double>(fitDatasCount);
            for (int i = 0; i < fitDatasCount; i++)
            {
                b[i] = Math.Log(IntensitySG[i]);
                a[i, 0] = 1;
                a[i, 1] = i;
                a[i, 2] = a[i, 1] * a[i, 1];
            }
            // Matrix.Equation(datas.Count, 3, a, b, X);
            //Matrix<double> m = Matrix<double>.Build.Random(3, 4);
            //Matrix<double> m = Matrix<double>.Build.Random(order, order+1);
            Matrix<double> matrixA = Matrix<double>.Build.DenseOfArray(a);
            // Matrix<double> matrixB = new MathNet.Numerics.LinearAlgebra.Matrix<double>(b, b.Length);
            Matrix<double> matrixB = Matrix<double>.Build.DenseOfColumnArrays(b);
            MathNet.Numerics.LinearAlgebra.Matrix<double> matrixC = matrixA.Solve(matrixB);
            for (int i = 0; i < order; i++)
            {
                X[i] = matrixC.ToArray()[i, 0];
            }
            //X = matrixC.SubMatrix
            double S = -1 / X[2];
            c = Math.Sqrt(-0.5 / X[2]);
            xMax = X[1] * S / 2.0; //最大值的index
            yMax = Math.Exp(X[0] + xMax * xMax / S);
            for (int i = 0; i < fitDatasCount; i++)
            {
                Gausian_y.Add(yMax * Math.Exp(-Math.Pow((i - xMax), 2) / S));
            }
            //  return Gausian_y;
        }


        private async static Task<bool> Gaussian_Task(MWArray PixelClip_M, MWArray IntensitySG_M, MWArray order_M) //完成後回傳TRUE
        {

            await Task.Run(() =>
            {
                GF gf = new GF();
                Sheet = gf.gaussfit(PixelClip_M, IntensitySG_M, order_M);


            });

            return true;
        }

        private static MWArray Sheet;

        public static void Gaussian_M(List<double> PixelClip_List, List<double> IntensitySG_List, double order, out List<double> Gausian_y_Result, out double xMax, out double yMax, out double sigma)
        {
            MWArray order_M = (MWNumericArray)order;
            MWArray PixelClip_M = (MWNumericArray)PixelClip_List.ToArray();
            MWArray IntensitySG_M = (MWNumericArray)IntensitySG_List.ToArray();


            /*     var processTask = Gaussian_Task(PixelClip_M, IntensitySG_M, order_M);
                 Task processFinishTask =Task.WhenAll(processTask);*/
            GF gf = new GF();
            Sheet = gf.gaussfit(PixelClip_M, IntensitySG_M, order_M);
            //  Task.Delay(500000);
            //   Thread.Sleep(500000);

            //      System.Threading.Thread.Sleep(10000);

            double[,] Sheet_Array;
            Sheet_Array = (double[,])((MWNumericArray)Sheet).ToArray();//try

            List<double> Gausian_y = new List<double>();
            for (int i = 0; i < IntensitySG_List.Count; i++)
            {

                Gausian_y.Add(Sheet_Array[0, i]);
            }
            Gausian_y_Result = Gausian_y;
            // Gausian_y = dd[0].ToList();
            sigma = Sheet_Array[1, 0];
            xMax = Sheet_Array[2, 0];
            yMax = Sheet_Array[3, 0];



        }
        public static Dictionary<string, List<double>> FindPeak_And_Gaussian(List<double> SG_Intensity, List<double> wavelength, double SlopeTh, double AmpTh, double SgWindowSize)
        {
            Dictionary<string, List<double>> Gaus_FWHL_Coef_Set = new Dictionary<string, List<double>>();
            PF Poly = new PF();
            FindPeaks fp = new FindPeaks();
            plot pt = new plot();
            PolyAndPlot PAP = new PolyAndPlot();

            MWArray IntensitySG_M = (MWNumericArray)SG_Intensity.ToArray();
            MWArray length = (MWNumericArray)SG_Intensity.Count;
            MWArray SlopeTreshold = (MWNumericArray)SlopeTh;
            MWArray AmpTreshold = (MWNumericArray)AmpTh;
            MWArray SG_WindowSize = (MWNumericArray)SgWindowSize;
            MWArray Peak_Group = (MWNumericArray)5;
            MWArray GaussianType = (MWNumericArray)3;

            MWArray yout_fp = fp.autofindpeaks(length, IntensitySG_M, SlopeTreshold, AmpTreshold, SG_WindowSize, Peak_Group, GaussianType);
            double[,] Peak;
            Peak = (double[,])((MWNumericArray)yout_fp).ToArray(MWArrayComponent.Real);
            List<double> Pixel = new List<double>();
            List<double> intensity = new List<double>();
            List<double> FWHL = new List<double>();

            for (int j = 1; j < Peak.GetUpperBound(1) + 1; j++)
            {
                for (int i = 0; i < Peak.GetUpperBound(0) + 1; i++)
                {
                    if (j == 1)
                    {
                        Pixel.Add(Peak[i, j]);
                    }
                    if (j == 2)
                    {
                        intensity.Add(Peak[i, j]);
                    }
                    if (j == 3)
                    {
                        FWHL.Add(Peak[i, j]);
                    }
                }
            }

            MWArray Pixel_M = (MWNumericArray)Pixel.ToArray();
            MWArray intensity_M = (MWNumericArray)intensity.ToArray();
            MWArray FWHL_M = (MWNumericArray)FWHL.ToArray();
            MWArray wavelength_M = (MWNumericArray)wavelength.ToArray();
            MWArray order_M = (MWNumericArray)3;
            List<double> FWHL_G = MWArray2Array2(FWHL_M, 8).ToList();

            //pt.Plot(Pixel_M, intensity_M, IntensitySG_M);

            if (Pixel.Count != wavelength.Count)
            {
                List<double> coef = new List<double>();
                pt.Plot(Pixel_M, intensity_M, IntensitySG_M);
                for (int i = 0; i < 4; i++)
                { coef.Add(0); }
                Gaus_FWHL_Coef_Set.Add("coef", coef);
                Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
                Gaus_FWHL_Coef_Set.Add("pixel_M", Pixel);
                return Gaus_FWHL_Coef_Set;
            }
            else
            {
                /*
                MWArray Fit = Poly.PolyFit(Pixel_M, wavelength_M, order_M);
                List<double> coef = MWArray2Array(Fit).ToList();
                Gaus_FWHL_Coef_Set.Add("coef", coef);
                Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
                Gaus_FWHL_Coef_Set.Add("pixel_M", Pixel);
                return Gaus_FWHL_Coef_Set;*/
                MWArray Fit = PAP.Poly_And_Plot(Pixel_M, wavelength_M, order_M, Pixel_M, intensity_M, IntensitySG_M);
                List<double> coef = MWArray2Array(Fit).ToList();
                Gaus_FWHL_Coef_Set.Add("coef", coef);
                Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
                Gaus_FWHL_Coef_Set.Add("pixel_M", Pixel);
                return Gaus_FWHL_Coef_Set;
            }
        }

        public static Dictionary<string, List<double>> Hg_FindPeak_And_Gaussian(List<double> Intensity, List<double> wavelength, double SlopeTh, double AmpTh, double SgWindowSize)
        {
            Dictionary<string, List<double>> Gaus_FWHL_Coef_Set = new Dictionary<string, List<double>>();
            PF Poly = new PF();
            FindPeaks fp = new FindPeaks();
            plot pt = new plot();
            PolyAndPlot PAP = new PolyAndPlot();

            MWArray IntensitySG_M = (MWNumericArray)Intensity.ToArray();
            MWArray length = (MWNumericArray)Intensity.Count;
            MWArray SlopeTreshold = (MWNumericArray)SlopeTh;
            MWArray AmpTreshold = (MWNumericArray)AmpTh;
            MWArray SG_WindowSize = (MWNumericArray)SgWindowSize;
            MWArray Peak_Group = (MWNumericArray)5;
            MWArray GaussianType = (MWNumericArray)3;

            MWArray yout_fp = fp.autofindpeaks(length, IntensitySG_M, SlopeTreshold, AmpTreshold, SG_WindowSize, Peak_Group, GaussianType);
            double[,] Peak;
            Peak = (double[,])((MWNumericArray)yout_fp).ToArray(MWArrayComponent.Real);
            List<double> Pixel = new List<double>();
            List<double> intensity = new List<double>();
            List<double> FWHL = new List<double>();

            for (int j = 1; j < Peak.GetUpperBound(1) + 1; j++)
            {
                for (int i = 0; i < Peak.GetUpperBound(0) + 1; i++)
                {
                    if (j == 1)
                    {
                        Pixel.Add(Peak[i, j]);
                    }
                    if (j == 2)
                    {
                        intensity.Add(Peak[i, j]);
                    }
                    if (j == 3)
                    {
                        FWHL.Add(Peak[i, j]);
                    }
                }
            }
            Pixel.RemoveRange(4, Pixel.Count - 4);
            intensity.RemoveRange(4, intensity.Count - 4);
            FWHL.RemoveRange(4, FWHL.Count - 4);
            MWArray FWHL_M = (MWNumericArray)FWHL.ToArray();
            List<double> FWHL_G = MWArray2Array2(FWHL_M, 8).ToList();

            Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
            Gaus_FWHL_Coef_Set.Add("pixel_M", Pixel);
            Gaus_FWHL_Coef_Set.Add("intensity_M", intensity);
            return Gaus_FWHL_Coef_Set;

        }


        public static Dictionary<string, List<double>> Ar_FindPeak_And_Gaussian(Dictionary<string, List<double>> FindPeak_Coef_Set, List<double> Pure_Intensity, List<double> Ar_Intensity, List<double> wavelength, double SlopeTh, double AmpTh, double SgWindowSize)
        {
            Dictionary<string, List<double>> Gaus_FWHL_Coef_Set = new Dictionary<string, List<double>>();
            PF Poly = new PF();
            FindPeaks fp = new FindPeaks();
            plot pt = new plot();
            PolyAndPlot PAP = new PolyAndPlot();

            MWArray IntensitySG_M = (MWNumericArray)Ar_Intensity.ToArray();
            MWArray length = (MWNumericArray)Ar_Intensity.Count;
            MWArray SlopeTreshold = (MWNumericArray)SlopeTh;
            MWArray AmpTreshold = (MWNumericArray)AmpTh;
            MWArray SG_WindowSize = (MWNumericArray)SgWindowSize;
            MWArray Peak_Group = (MWNumericArray)5;
            MWArray GaussianType = (MWNumericArray)3;

            MWArray yout_fp = fp.autofindpeaks(length, IntensitySG_M, SlopeTreshold, AmpTreshold, SG_WindowSize, Peak_Group, GaussianType);
            double[,] Peak;
            Peak = (double[,])((MWNumericArray)yout_fp).ToArray(MWArrayComponent.Real);

            Pure_Intensity.RemoveRange(600, Pure_Intensity.Count-600);
            for (int t = 600; t < Ar_Intensity.Count; t++)
            {
                Pure_Intensity.Add(Ar_Intensity[t]);
            }
            MWArray Intensity_for_plot = (MWNumericArray)Pure_Intensity.ToArray();
            bool first_find = false;
            bool second_find = false;
            FindPeak_Coef_Set["FWHL"].RemoveRange(4, FindPeak_Coef_Set["FWHL"].Count - 4);

            List<double> Pixel = new List<double>();
            List<double> intensity = new List<double>();
            List<double> FWHL = new List<double>();

            for (int j = 1; j < Peak.GetUpperBound(1) + 1; j++)
            {
                for (int i = 0; i < Peak.GetUpperBound(0) + 1; i++)
                {
                    if (j == 1)
                    {
                        Pixel.Add(Peak[i, j]);
                    }
                    if (j == 2)
                    {
                        intensity.Add(Peak[i, j]);
                    }
                    if (j == 3)
                    {
                        FWHL.Add(Peak[i, j]);
                    }
                }
            }
            for (int k = 0; k < Pixel.Count; k++)
            {
                if (FindPeak_Coef_Set["pixel_M"][3] + 220 <= Pixel[k] && Pixel[k] <= FindPeak_Coef_Set["pixel_M"][3] + 230)
                {
                    FindPeak_Coef_Set["intensity_M"].Add(intensity[k]);
                    FindPeak_Coef_Set["pixel_M"].Add(Pixel[k]);
                    FindPeak_Coef_Set["FWHL"].Add(FWHL[k]);
                    first_find = true;
                }
                else if (FindPeak_Coef_Set["pixel_M"][3] + 277 <= Pixel[k] && Pixel[k] <= FindPeak_Coef_Set["pixel_M"][3] + 287)
                {
                    FindPeak_Coef_Set["intensity_M"].Add(intensity[k]);
                    FindPeak_Coef_Set["pixel_M"].Add(Pixel[k]);
                    FindPeak_Coef_Set["FWHL"].Add(FWHL[k]);
                    second_find = true;
                }
            }
            if (first_find == false)
            {
                wavelength.RemoveAt(4);
            }
            else if (second_find == false)
            {
                wavelength.RemoveAt(5);
            }
            else if (first_find == false && second_find == false)
            {
                wavelength.RemoveRange(4, 2);
            }

                MWArray Pixel_M = (MWNumericArray)FindPeak_Coef_Set["pixel_M"].ToArray();
                MWArray intensity_M = (MWNumericArray)FindPeak_Coef_Set["intensity_M"].ToArray();
                MWArray FWHL_M = (MWNumericArray)FindPeak_Coef_Set["FWHL"].ToArray();
                MWArray wavelength_M = (MWNumericArray)wavelength.ToArray();
                MWArray order_M = (MWNumericArray)3;
                List<double> FWHL_G = MWArray2Array2(FWHL_M, 8).ToList();

                //pt.Plot(Pixel_M, intensity_M, IntensitySG_M);

                if (FindPeak_Coef_Set["pixel_M"].Count != wavelength.Count)
                {
                    List<double> coef = new List<double>();
                    pt.Plot(Pixel_M, intensity_M, Intensity_for_plot);
                    for (int i = 0; i < 4; i++)
                    { coef.Add(0); }
                    Gaus_FWHL_Coef_Set.Add("coef", coef);
                    Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
                    Gaus_FWHL_Coef_Set.Add("pixel_M", FindPeak_Coef_Set["pixel_M"]);
                    return Gaus_FWHL_Coef_Set;
                }
                else
                {
                    /*
                    MWArray Fit = Poly.PolyFit(Pixel_M, wavelength_M, order_M);
                    List<double> coef = MWArray2Array(Fit).ToList();
                    Gaus_FWHL_Coef_Set.Add("coef", coef);
                    Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
                    Gaus_FWHL_Coef_Set.Add("pixel_M", Pixel);
                    return Gaus_FWHL_Coef_Set;*/
                    MWArray Fit = PAP.Poly_And_Plot(Pixel_M, wavelength_M, order_M, Pixel_M, intensity_M, Intensity_for_plot);
                    List<double> coef = MWArray2Array(Fit).ToList();
                    Gaus_FWHL_Coef_Set.Add("coef", coef);
                    Gaus_FWHL_Coef_Set.Add("FWHL", FWHL_G);
                    Gaus_FWHL_Coef_Set.Add("pixel_M", FindPeak_Coef_Set["pixel_M"]);
                    return Gaus_FWHL_Coef_Set;
                }
            

        }
    }
}
