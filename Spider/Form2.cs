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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string head = HeadtextBox1.Text;
            //MessageBox.Show(head);
            string arm1 = Arm1textBox.Text;
            string arm2 = Arm2textBox.Text;
            string arm3 = Arm3textBox.Text;
            //string[] arm1Values = arm1.Split(',');
            //Console.WriteLine(arm1Values.Length);
            List<string> arms = new List<string>();
            arms.Add(arm1);
            arms.Add(arm2);
            arms.Add(arm3);
            foreach (string arm in arms)
            {
                int i = 1;
                string[] armValues = arm.Split(',');
                Console.WriteLine("Length of arm" + "is" + armValues.Length);
                i++;
            }
            DrawArms(arms);
        }
        public static void DrawArms(List<string> arms)
        {
            var visioApp = new visio.Application();
            var doc = visioApp.Documents.Add("");
            var page = visioApp.ActivePage;
            var shape6 = page.DrawOval(3.5, 6, 4.5, 5);
            shape6.Text = "Measure";
        }
    }
}
