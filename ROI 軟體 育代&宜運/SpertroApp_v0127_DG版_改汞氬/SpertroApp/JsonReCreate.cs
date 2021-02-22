using Newtonsoft.Json;
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

namespace SpertroApp
{
    public partial class JsonReCreate : Form
    {
        private static string Result = "{   }";
        string path = "";
        string ID = "";
        public JsonReCreate()
        {
            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog1.FileName;
            }
            string a = "";
             //Pass the file path and file name to the StreamReader constructor
             StreamReader sr = new StreamReader(path);
             //Read line of text
             while (!sr.EndOfStream)
             {
                textBox1.Text  += sr.ReadLine();
                textBox1.Text += "\r\n";
             }
             sr.Close();

                
        }

        private void JsonReCreate_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            string a = @"""";
            string b = ",";
            
            List<string> ROI = new List<string>();
            List<string> WL = new List<string>();
            //讀取JSON檔案
            using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
            {
                using (JsonTextReader reader = new JsonTextReader(sr))
                {
                    while (reader.Read())
                    {
                        if (reader.Value != null)
                        {
                            switch (reader.Value.ToString())
                            {
                                case "SC_ID":
                                    {
                                        reader.Read();
                                        ID = reader.Value.ToString();
                                    }
                                    break;

                                case "ID":
                                    {
                                        reader.Read();
                                        ID = reader.Value.ToString();
                                    }
                                    break;

                                case "ROI":
                                    {
                                        ROI.Clear();
                                        while (reader.Read())
                                        {
                                            if (reader.Value == null)
                                            {

                                                while (reader.Read())
                                                {
                                                    if (reader.Value == null)
                                                    {
                                                        break;
                                                    }
                                                    else if (reader.Value != null)
                                                    {
                                                        ROI.Add(reader.Value.ToString());
                                                    }
                                                }
                                                break;
                                            }
                                            else if (reader.Value != null)
                                            {
                                                ROI.Add(reader.Value.ToString());
                                                if (ROI[0].Split(',').Length == 4)
                                                {
                                                    string[] r = ROI[0].Split(',');
                                                    ROI.Clear();
                                                    foreach (string s in r)
                                                    {
                                                        ROI.Add(s);
                                                    }
                                                    break;
                                                }
                                            }
                                            else if (ROI.Count > 4)
                                            {
                                                break;
                                            }
                                        }

                                    }
                                    break;

                                case "WC_coefficients":
                                    WL.Clear();
                                    int k = 0;
                                    while (reader.Read())
                                    {
                                        if (reader.Value == null)
                                        {
                                            if (k == 1)
                                            { break; }
                                            k++;
                                        }
                                        else if (reader.Value != null)
                                        {
                                            WL.Add(reader.Value.ToString());
                                        }

                                    }
                                    break;

                                case "WL":
                                    WL.Clear();
                                    int t = 0;
                                    while (reader.Read())
                                    {
                                        if (reader.Value == null)
                                        {
                                            if (t == 1)
                                            { break; }
                                            t++;
                                        }
                                        else if (reader.Value != null)
                                        {
                                            WL.Add(reader.Value.ToString());
                                        }

                                    }
                                    break;
                            }
                        }
                    }
                    //讀取Int
                    //reader.ReadAsInt32().Value


                    //結束讀取
                    reader.Close();
                    sr.Close();
                }
            }
            Result =
            "{" + "\r\n" +
        a + "ID" + a + ":" + a + ID + a + b + "\r\n" +
        a + "ELC" + a + ":" + a + "5000" + a + b + "\r\n" +
        a + "GNV" + a + ":" + a + "32" + a + b + "\r\n" +
        a + "AGN" + a + ":" + a + "1X" + a + b + "\r\n" +
        a + "ROI" + a + ":" + a + ROI[0] + b + ROI[1] + b + ROI[2] + b + ROI[3] + a + b + "\r\n" +
        a + "WL" + a + ":" + "[" + a + WL[0] + a + b + a + WL[1] + a + b + a + WL[2] + a + b + a + WL[3] + a + "]" + "\r\n" +
           "}";
            textBox2.Text = Result;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path1 = @"json\isb" +"_" + ID+ "_"+"Ken" + ".json";
            string path = @"F:\isb.json";  //@"json\isb" + ".json";
            
            File.WriteAllText(path1, textBox2.Text);
            try
            {
                File.WriteAllText(path, textBox2.Text);
                MessageBox.Show("存檔完成");
            }
            catch
            {
                MessageBox.Show(@"一份已存至json\isb，另一份請確認E槽是否存在");
            }
        }
    }
}
