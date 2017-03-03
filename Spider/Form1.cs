using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using visio = Microsoft.Office.Interop.Visio;

namespace Spider
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        //event handlers
        private void showButton_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Load(openFileDialog1.FileName);
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
        }

        private void backgroundButton_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.BackColor = colorDialog1.Color;
            }
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            var visioApp = new visio.Application();
            var doc = visioApp.Documents.Add("");
            var page = visioApp.ActivePage;
            //var shape = page.DrawRectangle(1, 12, 2, 2);
            //var shape2 = page.DrawRectangle(5, 5, 5, 5);
            //var shape3 = page.DrawLine(1, 2, 5, 9);
            var shape5 = page.DrawCircularArc(4, 4,2, 1, 6.28);
            shape5.Text = "Measure";
            
            //var wordpp = new word.Application();
            //wordpp.Visible = true;
            //wordpp.Windows.Add();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
