using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpertroApp
{
    public partial class Get_SC_ID : Form
    {
        public Get_SC_ID()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.DialogResult = DialogResult.OK;
            this.Close();
        }


        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.DialogResult = DialogResult.OK;
                this.Close();

            }           
        }
    }
}
