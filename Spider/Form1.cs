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
            var line1 = page.DrawLine(4.125,5.5,7,5.5);
            var line2 = page.DrawLine(4, 5.5,6.25,3);
            var line3 = page.DrawLine(4, 5.5,4,2.5);
            var line4 = page.DrawLine(4, 5.5,1.5,3);
            var line5 = page.DrawLine(4, 5.5,1,5.5);
            var line6 = page.DrawLine(4, 5.5,1.5,8);
            var line7 = page.DrawLine(4, 5.5,6.5,8);
            var line8 = page.DrawLine(4, 5.5, 4, 8);
            //var shape5 = page.DrawCircularArc(4, 4,2,6.284,14);
            //shape5.Text = "Measure";
            var shape6 = page.DrawOval(3.5, 6, 4.5, 5);
            shape6.Text = "Measure";
            var shape7 = page.DrawOval(5.82, 5.62, 4.5, 5);
            //var wordpp = new word.Application();
            //wordpp.Visible = true;
            //wordpp.Windows.Add();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
