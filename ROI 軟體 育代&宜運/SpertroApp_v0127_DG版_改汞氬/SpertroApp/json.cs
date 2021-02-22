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
using System.Collections.Specialized;
using System.Web;
using System.Net;

namespace SpertroApp
{
    public partial class json : Form
    {
        private Form1 f1;
        
        public json(Form1 form)
        {
            InitializeComponent();
            f1 = form;
        }
        
        private void json_Load(object sender, EventArgs e)
        {
            string a = @"""";
            string b = ",";
            SC_ID.Text = f1.this_SC_ID;
            ROI.Text = f1.ROI_X.ToString() + "," + f1.ROI_W.ToString() + "," + f1.ROI_Y.ToString() + "," + f1.ROI_H.ToString();
            //-----------------------------------Wavelenght-----------------------------------------------
            for (int i = 0; i < f1.Wavelength.Count; i++)
            {
                if (i == f1.Wavelength.Count-1)
                {
                    Calibration_wavelengths.Text += a + f1.Wavelength[i] + a  + "]";
                    Calibration_pixels.Text += a + f1.pixel_M[i] + a + "]";
                    Calibration_after_wavelengths.Text += a + f1.Lamda_cal[i] + a + "]";
                    Calibration_wavelengths_errors_in_nm.Text += a + f1.DIFF[i] + a + "]";
                }
                else if (i == 0)
                { 
                    Calibration_wavelengths.Text += "[" + a + f1.Wavelength[i] + a + b;
                    Calibration_pixels.Text += "[" + a + f1.pixel_M[i] + a + b;
                    Calibration_after_wavelengths.Text += "[" + a + f1.Lamda_cal[i] + a + b;
                    Calibration_wavelengths_errors_in_nm.Text +="["+ a + f1.DIFF[i] + a + b;
                }
                else 
                { 
                    Calibration_wavelengths.Text += a + f1.Wavelength[i] + a + b;
                    Calibration_pixels.Text += a + f1.pixel_M[i] + a + b;
                    Calibration_after_wavelengths.Text += a + f1.Lamda_cal[i] + a + b;
                    Calibration_wavelengths_errors_in_nm.Text += a + f1.DIFF[i] + a + b;                  
                }
            }
            //--------------------------------------------------------------------------------------------
            Calibration_wavelengths_rms_errors_in_nm.Text = f1.DIFF_rms.ToString();
            RMS_FWHM_resolution_in_nm.Text = f1.LambdaSC_RMS.ToString();
            RMS_FWHM_spot_size_in_um.Text = f1.LambdaSC_Spot_RMS.ToString();
            RMS_FWHM_resolution_STD.Text = f1.LambdaSC_STD_100.ToString();
            RMS_FWHM_spot_size_STD.Text = f1.LambdaSC_Spot_STD_100.ToString();
            SNR_full_scale.Text = "["+a+f1.lamda_max_Intensity.ToString()+a+b+a+f1.SNR.ToString()+a+"]";
            Stray_light_in_nm.Text = f1.Stray_light.ToString();
            Dynamic_range.Text = f1.Dynamic_Range.ToString();
            Noise_rms_in_.Text = ((f1.RMS_of_Noise / 256) * 100).ToString();
            Noise_rms_in_256.Text = f1.RMS_of_Noise.ToString();
            try
            {
                Baseline_range_in_nm.Text = "[" + a + f1.Baseline_of_leftPoint.ToString() + a + b + a + f1.Baseline_of_rightPoint.ToString() + a + "]";
            }
            catch
            {
                MessageBox.Show("請先至'波形畫面'中計算出baseline_range!");
            }
            Digital_Gain.Text = f1.final_dg;
            Gamma.Text = f1.final_gamma;
            Backlight.Text = f1.final_back;
            SNR_of_White.Text = f1.White_Imax_Divided_By_DarkNoise.ToString();
            White_Max_Dark.Text = f1.White_Imax_Divided_By_Dark.ToString();
            try
            {
                HG_A0.Text = f1.Poly_Coefs_of_Hg_Ar[0].ToString();
                HG_A1.Text = f1.Poly_Coefs_of_Hg_Ar[1].ToString();
                HG_A2.Text = f1.Poly_Coefs_of_Hg_Ar[2].ToString();
                HG_A3.Text = f1.Poly_Coefs_of_Hg_Ar[3].ToString();
            }
            catch
            {
                MessageBox.Show("無汞氬燈量測數據");
            }
        }

        private void Fiber_to_SC_gap_in_mm_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path1 = @"json\isb" + "_" + SC_ID.Text + ".json";
            string path2 = @"F:\isb" + "_" + SC_ID.Text + ".json";
            File.WriteAllText(path1, textBox1.Text);
            try
            {
                File.WriteAllText(path2, textBox1.Text);
                MessageBox.Show("存檔完成");
            }
            catch
            {
                MessageBox.Show(@"一份已存至json\isb，另一份請確認E槽是否存在");
            }
            //上傳雲端
            string url = "http://corona-management.cloud.kahap.com/api/drive.php";         //巧禾API
            NameValueCollection postParams = HttpUtility.ParseQueryString(string.Empty);
            richTextBox1.Text = HttpUploadFile(url, path1, "file", "application / JSON", postParams);
            MessageBox.Show("上傳雲端已完成!");

            //-----------------------------------------------------------------------------------------------------
            string a = @"""";
            string b = ",";
            //---------------------------------------------拼寫json---------------------------------------
            string write =
                          "{" + "\r\n" +
                           a + "ID" + a + ":" + a + SC_ID.Text + a + b + "\r\n" +
                           a + "ELC" + a + ":" + a + "5000" + a + b + "\r\n" +
                           a + "GNV" + a + ":" + a + "32" + a + b + "\r\n" +
                           a + "AGN" + a + ":" + a + "1X" + a + b + "\r\n" +
                           a + "ROI" + a + ":" + a + f1.ROI_X.ToString() + b + f1.ROI_W.ToString() + b + f1.ROI_Y.ToString() + b + f1.ROI_H.ToString() + a + b + "\r\n" +
                           a + "WL" + a + ":" + "[" + a + f1.Poly_Coefs[0].ToString() + a + b + a + f1.Poly_Coefs[1].ToString() + a + b + a + f1.Poly_Coefs[2].ToString() + a + b + a + f1.Poly_Coefs[3].ToString() + a + "]" + "\r\n" +
                          "}";
            //--------------------------------------------------------------------------------------------
            //-----------------------------------------------存檔-----------------------------------------
            string path_ = @"json\isb" + "_" + f1.this_SC_ID + "_" + "Ken" + ".json";
            string path = @"F:\isb.json";  //@"json\isb" + ".json";

            File.WriteAllText(path_,write);
            try
            {
                File.WriteAllText(path, write);
                MessageBox.Show("存檔完成");
            }
            catch
            {
                MessageBox.Show(@"一份已存至json\isb，另一份請確認E槽是否存在");
            }
            //-------------------------------------------------------------------------------------------
        }
        public static string HttpUploadFile(string url, string file, string paramName, string contentType, NameValueCollection nvc)
        {
            //  Log.Debug(string.Format("Uploading {0} to {1}", file, url));
            string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundarybytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");
            string Resp;

            HttpWebRequest wr = (HttpWebRequest)WebRequest.Create(url);
            wr.ContentType = "multipart/form-data; boundary=" + boundary;
            wr.Method = "POST";
            wr.KeepAlive = true;
            wr.Credentials = System.Net.CredentialCache.DefaultCredentials;

            Stream rs = wr.GetRequestStream();

            string formdataTemplate = "Content-Disposition:multipart/form-data; name=\"{0}\"\r\n\r\n{1}";
            foreach (string key in nvc.Keys)
            {
                rs.Write(boundarybytes, 0, boundarybytes.Length);
                string formitem = string.Format(formdataTemplate, key, nvc[key]);
                byte[] formitembytes = System.Text.Encoding.UTF8.GetBytes(formitem);
                rs.Write(formitembytes, 0, formitembytes.Length);
            }
            rs.Write(boundarybytes, 0, boundarybytes.Length);

            string headerTemplate = "Content-Disposition: form-data; name=\"data\"; filename=\"{1}\"\r\nContent-Type: \"application/JSON\"\r\n\r\n";
            string header = string.Format(headerTemplate, paramName, file, contentType);
            byte[] headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
            rs.Write(headerbytes, 0, headerbytes.Length);

            FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[4096];
            int bytesRead = 0;
            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
            {
                rs.Write(buffer, 0, bytesRead);
            }
            fileStream.Close();

            byte[] trailer = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
            rs.Write(trailer, 0, trailer.Length);
            rs.Close();

            WebResponse wresp = null;
            try
            {
                wresp = wr.GetResponse();
                Stream stream2 = wresp.GetResponseStream();
                StreamReader reader2 = new StreamReader(stream2);
                // log.Debug(string.Format("File uploaded, server response is: {0}", reader2.ReadToEnd()));
                Resp = string.Format("JSON以上傳, 雲端伺服回應: {0}", reader2.ReadToEnd());
            }
            catch (Exception ex)
            {
                //log.Error("Error uploading file", ex);
                Resp = "上傳失敗";

                if (wresp != null)
                {
                    wresp.Close();
                    wresp = null;
                }
            }
            finally
            {
                wr = null;
            }

            return Resp;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            Date_of_calibration.Text = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
            string a = @"""";
            string b = ",";
            //--------------------------------------------------------------------------------------------------------------
            string ISBjson =
        "{"+"\r\n"+
        a+"ID"+a+":"+ a + SC_ID.Text + a+ b+"\r\n" +
        a + "ELC" + a + ":" + a + ELC.Text + a + b + "\r\n" +
        a + "GNV" + a + ":" + a + GNV.Text + a + b + "\r\n" +
        a + "AGN" + a + ":" + a + AGN.Text + a + b + "\r\n" +
        a + "ROI" + a + ":" + "[" + a + f1.ROI_X + a + b + a + f1.ROI_W + a + b + a + f1.ROI_Y + a + b + a + f1.ROI_H + a + "]" + b + "\r\n" +//UP
        a + "WL" + a + ":" + "[" + a + f1.Poly_Coefs[0] + a + b + a + f1.Poly_Coefs[1] + a + b + a + f1.Poly_Coefs[2] + a + b + a + f1.Poly_Coefs[3] + a + "]" + b + "\r\n" +//UP
        a +"Slit_width_in_um" +a+":" +a+ Slit_width_in_um.Text +a+b +"\r\n" +
        a+"Fiber_diameter_in_mm"+a+":" +a+ Fiber_diameter_in_mm.Text+a+b + "\r\n" +
        a+"Fiber_length_in_mm"+a+":" +a+ Fiber_length_in_mm.Text+a+b + "\r\n" +
        a+"Fiber_to_SC_gap_in_mm"+a+":"+ a + Fiber_to_SC_gap_in_mm.Text +a+b+ "\r\n" +
        a+"SC_dicing_version"+a+":"+a+ SC_dicing_version.Text+a+b+ "\r\n" +
        a+"SC_holder_version"+a+":"+a+ SC_holder_version.Text+a+b+ "\r\n" +
        a+"SC_top_wafer"+a+":"+a+ SC_top_wafer.Text+a+b+ "\r\n" +
        a+"SC_top_wafer_thickness_in_mm"+a+":"+a+ SC_top_wafer_thickness_in_mm.Text+a+b+ "\r\n" +
        a+"SC_bottom_wafer"+a+":"+a+ SC_bottom_wafer.Text+a+b+ "\r\n" +
        a+"SC_bottom_wafer_thickness_in_mm"+a+":"+a+ SC_bottom_wafer_thickness_in_mm.Text+a+b + "\r\n" +
        a+"IS_model"+a+":"+a+ IS_model.Text+a+b + "\r\n" +
        a+"IS_pixel_size_in_um"+a+":"+a+ IS_pixel_size_in_um.Text+a+b+ "\r\n" +
        a+"ISB_version"+a+":" +a+ ISB_version.Text+a+b + "\r\n" +
        a+"Date_of_calibration"+a+":" +a+ Date_of_calibration.Text +a+b+ "\r\n" +
        a+"Calibration_method"+a+":" +a+ Calibration_method.Text +a+b+ "\r\n" +
        a+"Calibration_wavelengths"+a+":" + Calibration_wavelengths.Text+b + "\r\n" +
        a+ "Calibration_pixels" + a + ":" + Calibration_pixels.Text + b + "\r\n" + //校正用Pixel
        a + "Calibration_after_wavelengths" + a + ":" + Calibration_after_wavelengths.Text + b + "\r\n" + //校正後波長
        a + "Calibration_wavelengths_errors_in_nm" + a + ":" + Calibration_wavelengths_errors_in_nm.Text + b + "\r\n" + //波長差異量
        a + "Calibration_wavelengths_rms_errors_in_nm" + a + ":" +a+ Calibration_wavelengths_rms_errors_in_nm.Text +a+ b + "\r\n" +//校正後誤差值量化        
        a +"RMS_FWHM_resolution_in_nm"+a+":"+a+ RMS_FWHM_resolution_in_nm.Text+a+b + "\r\n" +
        a+"RMS_FWHM_resolution_STD_in_%"+a+":"+a + RMS_FWHM_resolution_STD.Text+a+b + "\r\n" +
        a+"RMS_FWHM_spot_size_in_um"+a+":" +a+ RMS_FWHM_spot_size_in_um.Text+a+b+ "\r\n" +
        a+"RMS_FWHM_spot_size_STD_in_%"+a+":"+a+ RMS_FWHM_spot_size_STD.Text +a+b+ "\r\n"+
        a + "Spectral_accuracy_in_nm" + a + ":" + a + Spectral_accuracy_in_nm.Text + a + b + "\r\n" +
        a + "SNR_full_scale" + a + ":" + SNR_full_scale.Text + b + "\r\n" +
        a + "Stray_light_in_nm_%" + a + ":" + a + Stray_light_in_nm.Text + a + b + "\r\n" +
        a + "Dynamic_range" + a + ":" + a + Dynamic_range.Text + a +b+ "\r\n" +
        a + "Noise_rms_in_%" + a + ":" + a + Noise_rms_in_.Text + a +b+ "\r\n" +//雜訊佔最大值百分比
        a + "Noise_rms_in_256" + a + ":" + a + Noise_rms_in_256.Text + a +b+ "\r\n" +//雜訊值
        a + "Autoscaling_exposure_initial_value" + a + ":" + a + Autoscaling_exposure_initial_value.Text + a +b+ "\r\n" +//default AutoScaling_初始曝光時間
        a + "Baseline_range_in_nm" + a + ":" +Baseline_range_in_nm.Text +"\r\n"+//光譜基線範圍
        "}";
            //-------------------------------------------------------------------------------------------------------
            textBox1.AppendText(ISBjson);
        }

        private void SC_top_wafer_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            Date_of_calibration.Text = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
            string a = @"""";
            string b = ",";
            //--------------------------------------------------------------------------------------------------------------
            string ISBjson =
        "{" + "\r\n" +
        a + "ID" + a + ":" + a + SC_ID.Text + a + b + "\r\n" +
        a + "ELC" + a + ":" + a + ELC.Text + a + b + "\r\n" +
        a + "GNV" + a + ":" + a + GNV.Text + a + b + "\r\n" +
        a + "AGN" + a + ":" + a + AGN.Text + a + b + "\r\n" +
        a + "ROI" + a + ":" + "[" + a + f1.ROI_X + a + b + a + f1.ROI_W + a + b + a + f1.ROI_Y + a + b + a + f1.ROI_H + a + "]" + b + "\r\n" +//UP
        a + "WL" + a + ":" + "[" + a + f1.Poly_Coefs[0] + a + b + a + f1.Poly_Coefs[1] + a + b + a + f1.Poly_Coefs[2] + a + b + a + f1.Poly_Coefs[3] + a + "]" + b + "\r\n" +//UP
        a + "Slit_width_in_um" + a + ":" + a + Slit_width_in_um.Text + a + b + "\r\n" +
        a + "Fiber_diameter_in_mm" + a + ":" + a + Fiber_diameter_in_mm.Text + a + b + "\r\n" +
        a + "Fiber_length_in_mm" + a + ":" + a + Fiber_length_in_mm.Text + a + b + "\r\n" +
        a + "Fiber_to_SC_gap_in_mm" + a + ":" + a + Fiber_to_SC_gap_in_mm.Text + a + b + "\r\n" +
        a + "SC_dicing_version" + a + ":" + a + SC_dicing_version.Text + a + b + "\r\n" +
        a + "SC_holder_version" + a + ":" + a + SC_holder_version.Text + a + b + "\r\n" +
        a + "SC_top_wafer" + a + ":" + a + SC_top_wafer.Text + a + b + "\r\n" +
        a + "SC_top_wafer_thickness_in_mm" + a + ":" + a + SC_top_wafer_thickness_in_mm.Text + a + b + "\r\n" +
        a + "SC_bottom_wafer" + a + ":" + a + SC_bottom_wafer.Text + a + b + "\r\n" +
        a + "SC_bottom_wafer_thickness_in_mm" + a + ":" + a + SC_bottom_wafer_thickness_in_mm.Text + a + b + "\r\n" +
        a + "IS_model" + a + ":" + a + IS_model.Text + a + b + "\r\n" +
        a + "IS_pixel_size_in_um" + a + ":" + a + IS_pixel_size_in_um.Text + a + b + "\r\n" +
        a + "ISB_version" + a + ":" + a + ISB_version.Text + a + b + "\r\n" +
        a + "Date_of_calibration" + a + ":" + a + Date_of_calibration.Text + a + b + "\r\n" +
        a + "Calibration_method" + a + ":" + a + Calibration_method.Text + a + b + "\r\n" +
        a + "Calibration_wavelengths" + a + ":" + Calibration_wavelengths.Text + b + "\r\n" +
        a + "Calibration_pixels" + a + ":" + Calibration_pixels.Text + b + "\r\n" + //校正用Pixel
        a + "Calibration_after_wavelengths" + a + ":" + Calibration_after_wavelengths.Text + b + "\r\n" + //校正後波長
        a + "Calibration_wavelengths_errors_in_nm" + a + ":" + Calibration_wavelengths_errors_in_nm.Text + b + "\r\n" + //波長差異量
        a + "Calibration_wavelengths_rms_errors_in_nm" + a + ":" +a+Calibration_wavelengths_rms_errors_in_nm.Text+a+ b + "\r\n" +//校正後誤差值量化        
        a + "RMS_FWHM_resolution_in_nm" + a + ":" + a + RMS_FWHM_resolution_in_nm.Text + a + b + "\r\n" +
        a + "RMS_FWHM_resolution_STD_in_%" + a + ":" + a + RMS_FWHM_resolution_STD.Text + a + b + "\r\n" +
        a + "RMS_FWHM_spot_size_in_um" + a + ":" + a + RMS_FWHM_spot_size_in_um.Text + a + b + "\r\n" +
        a + "RMS_FWHM_spot_size_STD_in_%" + a + ":" + a + RMS_FWHM_spot_size_STD.Text + a + b + "\r\n" +
        a + "Spectral_accuracy_in_nm" + a + ":" + a + Spectral_accuracy_in_nm.Text + a + b + "\r\n" +
        a + "SNR_full_scale" + a + ":" + SNR_full_scale.Text + b + "\r\n" +
        a + "Stray_light_in_nm_%" + a + ":" + a + Stray_light_in_nm.Text + a + b + "\r\n" +
        a + "Dynamic_range" + a + ":" + a + Dynamic_range.Text + a + b + "\r\n" +
        a + "Noise_rms_in_%" + a + ":" + a + Noise_rms_in_.Text + a + b + "\r\n" +//雜訊佔最大值百分比
        a + "Noise_rms_in_256" + a + ":" + a + Noise_rms_in_256.Text + a + b + "\r\n" +//雜訊值
        a + "Autoscaling_exposure_initial_value" + a + ":" + a + Autoscaling_exposure_initial_value.Text + a + b + "\r\n" +//default AutoScaling_初始曝光時間
        a + "Baseline_range_in_nm" + a + ":" + Baseline_range_in_nm.Text + b + "\r\n" +//光譜基線範圍
        a + "Digital_Gain" + a + ":" + a + Digital_Gain.Text + a + b + "\r\n" +
        a + "Gamma" + a + ":" + a + Gamma.Text + a + b + "\r\n" +
        a + "Backlight" + a + ":" + a + Backlight.Text + a +b+ "\r\n" +
         a + "SNR_of_White" + a + ":" + a + SNR_of_White.Text + a + b + "\r\n" +
        a + "White_Max/Dark" + a + ":" + a + White_Max_Dark.Text + a + "\r\n" +
        "}";
            //-------------------------------------------------------------------------------------------------------
            textBox1.AppendText(ISBjson);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            Date_of_calibration.Text = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
            string a = @"""";
            string b = ",";
            //--------------------------------------------------------------------------------------------------------------
            string ISBjson =
        "{" + "\r\n" +
        a + "ID" + a + ":" + a + SC_ID.Text + a + b + "\r\n" +
        a + "ELC" + a + ":" + a + ELC.Text + a + b + "\r\n" +
        a + "GNV" + a + ":" + a + GNV.Text + a + b + "\r\n" +
        a + "AGN" + a + ":" + a + AGN.Text + a + b + "\r\n" +
        a + "ROI" + a + ":" + "[" + a + f1.ROI_X + a + b + a + f1.ROI_W + a + b + a + f1.ROI_Y + a + b + a + f1.ROI_H + a + "]" + b + "\r\n" +//UP
        a + "WL" + a + ":" + "[" + a + f1.Poly_Coefs_of_Hg_Ar[0].ToString() + a + b + a + f1.Poly_Coefs_of_Hg_Ar[1].ToString() + a + b + a + f1.Poly_Coefs_of_Hg_Ar[2].ToString() + a + b + a + f1.Poly_Coefs_of_Hg_Ar[3].ToString() + a + "]" + b + "\r\n" +//UP
        a + "Slit_width_in_um" + a + ":" + a + Slit_width_in_um.Text + a + b + "\r\n" +
        a + "Fiber_diameter_in_mm" + a + ":" + a + Fiber_diameter_in_mm.Text + a + b + "\r\n" +
        a + "Fiber_length_in_mm" + a + ":" + a + Fiber_length_in_mm.Text + a + b + "\r\n" +
        a + "Fiber_to_SC_gap_in_mm" + a + ":" + a + Fiber_to_SC_gap_in_mm.Text + a + b + "\r\n" +
        a + "SC_dicing_version" + a + ":" + a + SC_dicing_version.Text + a + b + "\r\n" +
        a + "SC_holder_version" + a + ":" + a + SC_holder_version.Text + a + b + "\r\n" +
        a + "SC_top_wafer" + a + ":" + a + SC_top_wafer.Text + a + b + "\r\n" +
        a + "SC_top_wafer_thickness_in_mm" + a + ":" + a + SC_top_wafer_thickness_in_mm.Text + a + b + "\r\n" +
        a + "SC_bottom_wafer" + a + ":" + a + SC_bottom_wafer.Text + a + b + "\r\n" +
        a + "SC_bottom_wafer_thickness_in_mm" + a + ":" + a + SC_bottom_wafer_thickness_in_mm.Text + a + b + "\r\n" +
        a + "IS_model" + a + ":" + a + IS_model.Text + a + b + "\r\n" +
        a + "IS_pixel_size_in_um" + a + ":" + a + IS_pixel_size_in_um.Text + a + b + "\r\n" +
        a + "ISB_version" + a + ":" + a + ISB_version.Text + a + b + "\r\n" +
        a + "Date_of_calibration" + a + ":" + a + Date_of_calibration.Text + a + b + "\r\n" +
        a + "Calibration_method" + a + ":" + a + Calibration_method.Text + a + b + "\r\n" +
        a + "Calibration_wavelengths" + a + ":" + Calibration_wavelengths.Text + b + "\r\n" +
        a + "Calibration_pixels" + a + ":" + Calibration_pixels.Text + b + "\r\n" + //校正用Pixel
        a + "Calibration_after_wavelengths" + a + ":" + Calibration_after_wavelengths.Text + b + "\r\n" + //校正後波長
        a + "Calibration_wavelengths_errors_in_nm" + a + ":" + Calibration_wavelengths_errors_in_nm.Text + b + "\r\n" + //波長差異量
        a + "Calibration_wavelengths_rms_errors_in_nm" + a + ":" + a + Calibration_wavelengths_rms_errors_in_nm.Text + a + b + "\r\n" +//校正後誤差值量化        
        a + "RMS_FWHM_resolution_in_nm" + a + ":" + a + RMS_FWHM_resolution_in_nm.Text + a + b + "\r\n" +
        a + "RMS_FWHM_resolution_STD_in_%" + a + ":" + a + RMS_FWHM_resolution_STD.Text + a + b + "\r\n" +
        a + "RMS_FWHM_spot_size_in_um" + a + ":" + a + RMS_FWHM_spot_size_in_um.Text + a + b + "\r\n" +
        a + "RMS_FWHM_spot_size_STD_in_%" + a + ":" + a + RMS_FWHM_spot_size_STD.Text + a + b + "\r\n" +
        a + "Spectral_accuracy_in_nm" + a + ":" + a + Spectral_accuracy_in_nm.Text + a + b + "\r\n" +
        a + "SNR_full_scale" + a + ":" + SNR_full_scale.Text + b + "\r\n" +
        a + "Stray_light_in_nm_%" + a + ":" + a + Stray_light_in_nm.Text + a + b + "\r\n" +
        a + "Dynamic_range" + a + ":" + a + Dynamic_range.Text + a + b + "\r\n" +
        a + "Noise_rms_in_%" + a + ":" + a + Noise_rms_in_.Text + a + b + "\r\n" +//雜訊佔最大值百分比
        a + "Noise_rms_in_256" + a + ":" + a + Noise_rms_in_256.Text + a + b + "\r\n" +//雜訊值
        a + "Autoscaling_exposure_initial_value" + a + ":" + a + Autoscaling_exposure_initial_value.Text + a + b + "\r\n" +//default AutoScaling_初始曝光時間
        a + "Baseline_range_in_nm" + a + ":" + Baseline_range_in_nm.Text + b + "\r\n" +//光譜基線範圍
        a + "Digital_Gain" + a + ":" + a + Digital_Gain.Text + a + b + "\r\n" +
        a + "Gamma" + a + ":" + a + Gamma.Text + a + b + "\r\n" +
        a + "Backlight" + a + ":" + a + Backlight.Text + a + b + "\r\n" +
         a + "SNR_of_White" + a + ":" + a + SNR_of_White.Text + a + b + "\r\n" +
        a + "White_Max/Dark" + a + ":" + a + White_Max_Dark.Text + a + "\r\n" +
        "}";
            //-------------------------------------------------------------------------------------------------------
            textBox1.AppendText(ISBjson);
        }
    }
}
