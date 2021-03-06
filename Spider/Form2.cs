﻿using System;
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
                int limbLength = 6;
                double xEndIncrement = limbLength * Math.Cos(theta);
                double yEndIncrement = limbLength * Math.Sin(theta);

                Coordinate limbStart = new Coordinate(origin.x + xStartIncrement, origin.y + yStartIncrement);
                Coordinate limbEnd = new Coordinate(limbStart.x + xEndIncrement, limbStart.y + yEndIncrement);
                
                var limb = page.DrawLine(limbStart.x,limbStart.y,limbEnd.x, limbEnd.y);
                List<Spot> spots = limbs[i].spots;
                List<Coordinate> pointsBetween = getPointsBetween(limbStart.x, limbStart.y, limbEnd.x, limbEnd.y,spots.Count);
                for (int j = 0; j < pointsBetween.Count;j++ )
                {
                   //draw the spots
                    double spotRadius = .15;
                    //var cicleArc = page.DrawCircularArc(pointsBetween[j].x, pointsBetween[j].y, .15, 0, 2*Math.PI);
                    var spotDrawing = page.DrawOval(pointsBetween[j].x - spotRadius, pointsBetween[j].y + spotRadius, pointsBetween[j].x + spotRadius, pointsBetween[j].y - spotRadius);
                    spotDrawing.Text = spots[j].spotText;
                }
                theta += thetaIncrement;

            }

            headCircle.Text = spidey.head.headText;
        }
        public List<Coordinate> getPointsBetween(double x1, double y1, double x2, double y2, int spotCount)
        {
            List<Coordinate> pointsBetween = new List<Coordinate>();
            //y = mx+c
            if (x1 != x2)
            {
                double m = (y2 - y1) / (x2 - x1);
                double b = y1 - m * x1;
                double deltaX = (x2 - x1) / spotCount;
                for (int i = 1; i <= spotCount; i++)
                {
                    Coordinate point = new Coordinate();
                    point.x = x1 + (deltaX * i);
                    point.y = m * (point.x) + b;
                    pointsBetween.Add(point);
                }
            }
            else
            {
                double deltaY = (y2 - y1) / spotCount;
                for (int i = 1; i <= spotCount; i++)
                {
                    Coordinate point = new Coordinate();
                    point.x = x1;
                    point.y = y1 + (deltaY * i);
                    pointsBetween.Add(point);
                }
            }
            return pointsBetween; 
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
        public Coordinate()
        {
            
        }

    }
}

//determine the lines bases on the hours hand on the clock
//2 arms --> 9---*---3
//3 arms --> 11---7---3
//4 arms --> 12---9---6---3
//5 arms --> 10---7---5---2---12
//6 arms --> 10---12---2---4---6---8

//var cicleArc = page.DrawCircularArc((limbStart.x + limbEnd.x) / 2, (limbStart.y + limbEnd.y) / 2, .15, 6.284, 14);
                  //  var cicleArc2 = page.DrawCircularArc(limbEnd.x, limbEnd.y, .15, 6.284, 14);