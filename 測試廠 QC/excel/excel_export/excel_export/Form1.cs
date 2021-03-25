using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace excel_export
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\yuchi\source\repos\excel_export\excel_export\bin\Debug\Report";
            //----------
            var Report_Data = new List<QCreport>
            {
                new QCreport
                {
                               Date = DateTime.Now.ToString("yyyy_MM_dd"),
                               OP_ID = "M10812015",
                               Mechine_ID = "V6-1205",
                               FW_ver = "v1.27",
                               Chip_ID ="20b4-105-c3",
                               Initial_Data =new string[] {"0.7","2.0","3.0"},
                               WC_coef = new string[]{ "-0.123456878", "0.123456878", "0.123456878", "0.123456878*E-8" },
                               Image_Date=new string[]{ "roi_x", "roi_w","roi_y","roi_h", "exp","dg","ag"},
                               Motor_Data=new string[]{ "motor", "motor_err" },
                               LED_A=new string[]{ "LED_A", "LEDA_STD" },
                               LED_B=new string[]{ "LED_B", "LEDB_STD" },
                               Spectrum_A=new string[]{ "SpectrumA", "SpectrumA_STD" },
                               Spectrum_B=new string[]{ "SpectrumB", "SpectrumB_STD" }
        }
            };
            if (File.Exists(path + @"\" + "Report_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx") == false)
            {
                Export_Excel.CreatExcel(path +@"\"+"Report_" + DateTime.Now.ToString("yyyy_MM_dd")+".xlsx","QC");
            }
            Export_Excel.WriteToExcel( path+ @"\" + "Report_" + DateTime.Now.ToString("yyyy_MM_dd") , Report_Data);
            MessageBox.Show("存檔完成");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\yuchi\source\repos\excel_export\excel_export\bin\Debug\Report";
            //----------
            var BLE_Report_Data = new List<BLEreport>
            {
                new BLEreport
                {
                               Date = DateTime.Now.ToString("yyyy_MM_dd"),
                               OP_ID = "M10812015",
                               Mechine_ID = "V6-1205",
                               Chip_ID ="20b4-105-c3",
                               BLE = "BLeeeeeee"
        }
            };
            if (File.Exists(path + @"\" + "BLE_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx") == false)
            {
                Export_Excel.CreatExcel(path + @"\" + "BLE_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx", "BLE");
            }
            Export_Excel.WriteTo_BLE_Excel(path + @"\" + "BLE_" + DateTime.Now.ToString("yyyy_MM_dd"), BLE_Report_Data);
            MessageBox.Show("存檔完成");
        }
    }
    
}
