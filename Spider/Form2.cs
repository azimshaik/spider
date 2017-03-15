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
            string headVal = HeadtextBox1.Text;
            List<string> textBoxes = new List<string>();
            string textBox1 = Arm1textBox.Text;
            string textBox2 = Arm2textBox.Text;
            string textBox3 = Arm3textBox.Text;
            string textBox4 = Arm4textBox.Text;
            string textBox5 = Arm5textBox.Text;
            string textBox6 = Arm6textBox.Text;
            string textBox7 = Arm7textBox.Text;
            string textBox8 = Arm8textBox.Text;
            textBoxes.Add(textBox1);
            textBoxes.Add(textBox2);
            textBoxes.Add(textBox3);
            textBoxes.Add(textBox4);
            textBoxes.Add(textBox5);
            textBoxes.Add(textBox6);
            textBoxes.Add(textBox7);
            textBoxes.Add(textBox8);
            Spider spidey = new Spider();
            //Spot spot = new Spot();
            Head head = new Head();
            head.headText = headVal;
            List<Limb> limbs = new List<Limb>();
            foreach(string textBox in textBoxes){
                List<Spot> spots = new List<Spot>();
                string[] spotvals = textBox.Split(',');
                foreach(string spotval in spotvals)
                {
                    Spot spot = new Spot();
                    spot.spotText = spotval;
                    spots.Add(spot);
                }
                Limb limb = new Limb();
                limb.spots = spots;
                limbs.Add(limb);
            }
            spidey.limbs = limbs;
            spidey.head = head;
            drawSpider(spidey);
        }
        private void drawSpider(Spider spidey)
        {
            var visioApp = new visio.Application();
            var doc = visioApp.Documents.Add("");
            var page = visioApp.ActivePage;

            Coordinate topLeft = new Coordinate(3.5,6);
            Coordinate bottomRight = new Coordinate(4.5, 5);
            var headCircle = page.DrawOval(topLeft.x, topLeft.y, bottomRight.x, bottomRight.y);
            Coordinate origin = new Coordinate((topLeft.x + bottomRight.x) / 2, (topLeft.y + bottomRight.y) / 2);
            double radius = Math.Abs(topLeft.x - bottomRight.x) / 2;
            List<Limb> limbs = spidey.limbs;
            double thetaIncrement = 2*Math.PI / limbs.Count;
            double theta = 0;
            for (int i = 0; i < limbs.Count; i++)
            {
                double xStartIncrement = radius * Math.Cos(theta);
                double yStartIncrement = radius * Math.Sin(theta);

                //HARDCODED
                int limbLength = 5;
                double xEndIncrement = limbLength * Math.Cos(theta);
                double yEndIncrement = limbLength * Math.Sin(theta);

                Coordinate limbStart = new Coordinate(origin.x + xStartIncrement, origin.y + yStartIncrement);
                Coordinate limbEnd = new Coordinate(limbStart.x + xEndIncrement, limbStart.y + yEndIncrement);
                
                var limb = page.DrawLine(limbStart.x,limbStart.y,limbEnd.x, limbEnd.y);
                theta += thetaIncrement;

            }

            headCircle.Text = spidey.head.headText;
        }
    
    }
    public class Spider
    {
        public Head head {get; set;}
        public List<Limb> limbs{get; set;}
    }
    public class Head
    {
        public string headText{ get; set;}
    }
    public class Limb
    {
        public List<Spot> spots { get; set; }
    }
    public class Spot
    {
        public string spotText { get; set; }
    }
    public class Line
    {
        public Coordinate startPoint { get; set; }
        public Coordinate endPoint {get; set;}
    }
    public class Circle
    {
        public Coordinate cornerTopLeft { get; set; }
        public Coordinate cornerBottomRight { get; set; }
    }
    public class Coordinate
    {
        public double x {get; set;}
        public double y {get; set;}
        public Coordinate(double x, double y)
        {
            this.x = x;
            this.y = y;
        }

    }
}

//determine the lines bases on the hours hand on the clock
//2 arms --> 9---*---3
//3 arms --> 11---7---3
//4 arms --> 12---9---6---3
//5 arms --> 10---7---5---2---12
//6 arms --> 10---12---2---4---6---8
/*
var visioApp = new visio.Application();
            var doc = visioApp.Documents.Add("");
            var page = visioApp.ActivePage;
            var shape6 = page.DrawOval(3.5, 6, 4.5, 5);// Center (4, 5.5)
            shape6.Text = "Measure";
            var line1 = page.DrawLine(4, 5.5, 7, 5.5);
    if (!string.IsNullOrWhiteSpace(arm2))
            {
                arms.Add(arm2);
            }
 
 //string[] arm1Values = arm1.Split(',');
 */