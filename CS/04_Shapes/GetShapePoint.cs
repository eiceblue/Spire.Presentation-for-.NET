using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetShapePoint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"../../../../../../Data/ShapePoint.pptx");

            //Get the first shape in first slide
            IAutoShape shape = (IAutoShape)ppt.Slides[0].Shapes[0];

            //Get the Point of shape
            IList <PointF> points = shape.Points;

            StringBuilder sb = new StringBuilder();
            sb.Append("point count£º" + " " + points.Count + "\r\n");

            for (int i = 0; i < points.Count; i++)
            {
                sb.Append("point" + i + " " + points[i] + "\r\n");
            }

            //Save the result txt file           
            File.WriteAllText("PointInformation.txt", sb.ToString());
            System.Diagnostics.Process.Start("PointInformation.txt");
        }
    }
}