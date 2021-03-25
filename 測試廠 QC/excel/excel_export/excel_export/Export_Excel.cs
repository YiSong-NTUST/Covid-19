using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_export
{
    public class QCreport
    {
        public string Date{ get; set; }
        public string OP_ID{ get; set; }
        public string Mechine_ID { get; set; }
        public string FW_ver { get; set; }
        public string Chip_ID{ get; set; }  
        public string[] Initial_Data{ get; set; }//Initial_Data[Xts,T1,T2]
        public string[] WC_coef { get; set; } //WC_coef[A0,A1,A2,A3]
        public string[] Image_Date { get; set; }//[ROI,ROI,ROI,ROI,EXP,DG,AG]
        public string[] Motor_Data { get; set; }//[motor,motor精度誤差]
        public string[] LED_A { get; set; }//[LED_A,LED_A STD]
        public string[] LED_B { get; set; }//[LED_B,LED_B STD]
        public string[] Spectrum_A { get; set; }//[Spectrum_A,Spectrum_A STD]
        public string[] Spectrum_B { get; set; }//[Spectrum_B,Spectrum_B STD]

    }
    public class BLEreport
    {
        public string Date { get; set; }
        public string OP_ID { get; set; }
        public string Mechine_ID { get; set; }
        public string Chip_ID { get; set; }
        public string BLE { get; set; }
    }

    class Export_Excel
    {

        public static void CreatExcel(string path, string which)
        {
            //create
            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Sheets[1];

            if (which == "BLE")
            {
                //==========================EXCEL2=========================================

                worksheet.Name = "Report";
                #region 第一行
                worksheet.Cells[1, "A"] = "日期";
                worksheet.Cells[1, "B"] = "工號";
                worksheet.Cells[1, "C"] = "機台編號";
                worksheet.Cells[1, "D"] = "晶片編號";
                worksheet.Cells[1, "E"] = "藍芽";
                #endregion
                #region 第二行
                worksheet.Cells[2, "A"] = "Date";
                worksheet.Cells[2, "B"] = "OP ID";
                worksheet.Cells[2, "C"] = "Machine ID";
                worksheet.Cells[2, "D"] = "Chip ID";
                worksheet.Cells[2, "E"] = "BLE";
                #endregion
            }
            else if(which == "QC")
            {
                //===========================EXCEL1========================================
                worksheet.Name = "Report";
                #region 第一行
                worksheet.Cells[1, "A"] = "日期";
                worksheet.Cells[1, "B"] = "工號";
                worksheet.Cells[1, "C"] = "機台編號";
                worksheet.Cells[1, "D"] = "韌體版本";
                worksheet.Cells[1, "E"] = "晶片編號";
                //=========================================================================
                Excel.Range range = worksheet.Range[worksheet.Cells[1, "F"], worksheet.Cells[1, "H"]];
                range.Merge(true);
                range.Value2 = "初始資訊";
                Excel.Range range1 = worksheet.Range[worksheet.Cells[1, "I"], worksheet.Cells[1, "L"]];
                range1.Merge(true);
                range1.Value2 = "波長係數";
                Excel.Range range2 = worksheet.Range[worksheet.Cells[1, "M"], worksheet.Cells[1, "S"]];
                range2.Merge(true);
                range2.Value2 = "影像資訊";
                Excel.Range range3 = worksheet.Range[worksheet.Cells[1, "T"], worksheet.Cells[1, "U"]];
                range3.Merge(true);
                range3.Value2 = "馬達誤差";
                Excel.Range range4 = worksheet.Range[worksheet.Cells[1, "V"], worksheet.Cells[1, "W"]];
                range4.Merge(true);
                range4.Value2 = "LED A燈資訊";
                Excel.Range range5 = worksheet.Range[worksheet.Cells[1, "X"], worksheet.Cells[1, "Y"]];
                range5.Merge(true);
                range5.Value2 = "LED B燈資訊";
                Excel.Range range6 = worksheet.Range[worksheet.Cells[1, "Z"], worksheet.Cells[1, "AA"]];
                range6.Merge(true);
                range6.Value2 = "光譜 A燈資訊";
                Excel.Range range7 = worksheet.Range[worksheet.Cells[1, "AB"], worksheet.Cells[1, "AC"]];
                range7.Merge(true);
                range7.Value2 = "光譜 B燈資訊";
                #endregion
                #region 第二行
                worksheet.Cells[2, "A"] = "Date";
                worksheet.Cells[2, "B"] = "OP ID";
                worksheet.Cells[2, "C"] = "Machine ID";
                worksheet.Cells[2, "D"] = "FW Ver.";
                worksheet.Cells[2, "E"] = "Chip ID";
                worksheet.Cells[2, "F"] = "Xts";
                worksheet.Cells[2, "G"] = "T1";
                worksheet.Cells[2, "H"] = "T2";
                worksheet.Cells[2, "I"] = "A0";
                worksheet.Cells[2, "J"] = "A1";
                worksheet.Cells[2, "K"] = "A2";
                worksheet.Cells[2, "L"] = "A3";
                worksheet.Cells[2, "M"] = "ROI";
                worksheet.Cells[2, "N"] = "ROI";
                worksheet.Cells[2, "O"] = "ROI";
                worksheet.Cells[2, "P"] = "ROI";
                worksheet.Cells[2, "Q"] = "EXP";
                worksheet.Cells[2, "R"] = "DG";
                worksheet.Cells[2, "S"] = "AG";
                worksheet.Cells[2, "T"] = "Motor";
                worksheet.Cells[2, "U"] = "Motor 精度誤差";
                worksheet.Cells[2, "V"] = "LED-A";
                worksheet.Cells[2, "W"] = "LED-A STD";
                worksheet.Cells[2, "X"] = "LED-B";
                worksheet.Cells[2, "Y"] = "LED-B STD";
                worksheet.Cells[2, "Z"] = "Spectrum-A";
                worksheet.Cells[2, "AA"] = "Spectrum-A STD";
                worksheet.Cells[2, "AB"] = "Spectrum-B";
                worksheet.Cells[2, "AC"] = "Spectrum-B STD";
                #endregion
            }
            //=========================================================================
            worksheet.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();

        }
        /// <summary>
        /// open an excel file,then write the content to file
        /// </summary>
        /// <param name="FileName">file name</param>
        /// <param name="findString">first cloumn</param>
        /// <param name="replaceString">second cloumn</param>
        public static void WriteToExcel(string excelName, IEnumerable<QCreport> QCReport_data)
        {
            //open
            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook mybook = app.Workbooks.Open(excelName, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
            Excel.Worksheet mysheet = (Excel.Worksheet)mybook.Worksheets[1];
            mysheet.Activate();
            //get activate sheet max row count
            int maxrow = mysheet.UsedRange.Rows.Count + 1;
            int maxcolumn = mysheet.UsedRange.Columns.Count + 1;
            foreach (var data in QCReport_data)
            {
                mysheet.Cells[maxrow, "A"] = data.Date;
                mysheet.Cells[maxrow, "B"] = data.OP_ID;
                mysheet.Cells[maxrow, "C"] = data.Mechine_ID;
                mysheet.Cells[maxrow, "D"] = data.FW_ver;
                mysheet.Cells[maxrow, "E"] = data.Chip_ID;
                mysheet.Cells[maxrow, "F"] = data.Initial_Data[0];
                mysheet.Cells[maxrow, "G"] = data.Initial_Data[1];
                mysheet.Cells[maxrow, "H"] = data.Initial_Data[2];
                mysheet.Cells[maxrow, "I"] = data.WC_coef[0];
                mysheet.Cells[maxrow, "J"] = data.WC_coef[1];
                mysheet.Cells[maxrow, "K"] = data.WC_coef[2];
                mysheet.Cells[maxrow, "L"] = data.WC_coef[3];
                mysheet.Cells[maxrow, "M"] = data.Image_Date[0];
                mysheet.Cells[maxrow, "N"] = data.Image_Date[1];
                mysheet.Cells[maxrow, "O"] = data.Image_Date[2];
                mysheet.Cells[maxrow, "P"] = data.Image_Date[3];
                mysheet.Cells[maxrow, "Q"] = data.Image_Date[4];
                mysheet.Cells[maxrow, "R"] = data.Image_Date[5];
                mysheet.Cells[maxrow, "S"] = data.Image_Date[6];
                mysheet.Cells[maxrow, "T"] = data.Motor_Data[0];
                mysheet.Cells[maxrow, "U"] = data.Motor_Data[1];
                mysheet.Cells[maxrow, "V"] = data.LED_A[0];
                mysheet.Cells[maxrow, "W"] = data.LED_A[1];
                mysheet.Cells[maxrow, "X"] = data.LED_B[0];
                mysheet.Cells[maxrow, "Y"] = data.LED_B[1];
                mysheet.Cells[maxrow, "Z"] = data.Spectrum_A[0];
                mysheet.Cells[maxrow, "AA"] = data.Spectrum_A[1];
                mysheet.Cells[maxrow, "AB"] = data.Spectrum_B[0];
                mysheet.Cells[maxrow, "AC"] = data.Spectrum_B[1];
            }
            //============================================================
            for (int i = 0; i < maxcolumn; i++)
            {
                for (int j = 0; j < maxrow; j++)
                {
                    Excel.Range contentRange = mysheet.Cells[j + 1, i + 1] as Excel.Range;//不用標題則從第二行開始

                    contentRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中

                    contentRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中
                    contentRange.Font.Size = 15;
                    contentRange.Font.Bold = true;
                    contentRange.EntireColumn.AutoFit(); //自動調整列寬
                    //contentRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//設定邊框
                    // contentRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//邊框常規粗細
                    contentRange.WrapText = false;//自動換行

                    contentRange.NumberFormatLocal = "@";//文字格式
                }
            }
            File.Delete(excelName);
            mybook.Save();
            mybook.Close(false, Type.Missing, Type.Missing);
            mybook = null;
            //quit excel app
            app.Quit();
            System.Diagnostics.Process[] excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (var item in excelProcess)
            {
                item.Kill();
            }
        }
        public static void WriteTo_BLE_Excel(string excelName, IEnumerable<BLEreport> BLE_Report_data)
        {
            //open
            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook mybook = app.Workbooks.Open(excelName, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
            Excel.Worksheet mysheet = (Excel.Worksheet)mybook.Worksheets[1];
            mysheet.Activate();
            //get activate sheet max row count
            int maxrow = mysheet.UsedRange.Rows.Count + 1;
            int maxcolumn = mysheet.UsedRange.Columns.Count + 1;
            foreach (var data in BLE_Report_data)
            {
                mysheet.Cells[maxrow, "A"] = data.Date;
                mysheet.Cells[maxrow, "B"] = data.OP_ID;
                mysheet.Cells[maxrow, "C"] = data.Mechine_ID;
                mysheet.Cells[maxrow, "D"] = data.Chip_ID;
                mysheet.Cells[maxrow, "E"] = data.BLE;
            }
            //============================================================
            for (int i = 0; i < maxcolumn; i++)
            {
                for (int j = 0; j < maxrow; j++)
                {
                    Excel.Range contentRange = mysheet.Cells[j + 1, i + 1] as Excel.Range;//不用標題則從第二行開始

                    contentRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//水平居中

                    contentRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;//垂直居中
                    contentRange.Font.Size = 15;
                    contentRange.Font.Bold = true;
                    contentRange.EntireColumn.AutoFit(); //自動調整列寬
                    //contentRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//設定邊框
                    // contentRange.Borders.Weight = Excel.XlBorderWeight.xlThin;//邊框常規粗細
                    contentRange.WrapText = false;//自動換行

                    contentRange.NumberFormatLocal = "@";//文字格式
                }
            }
            File.Delete(excelName);
            mybook.Save();
            mybook.Close(false, Type.Missing, Type.Missing);
            mybook = null;
            //quit excel app
            app.Quit();
            System.Diagnostics.Process[] excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (var item in excelProcess)
            {
                item.Kill();
            }
        }
    }
}
