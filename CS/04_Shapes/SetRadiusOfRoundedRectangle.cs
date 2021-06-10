using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;


namespace SetRadiusOfRoundedRectangle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Insert a rounded rectangle and set its radious
            presentation.Slides[0].Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10);

            //Append a rounded rectangle and set its radius
            IAutoShape shape = presentation.Slides[0].Shapes.AppendRoundRectangle(380, 180, 100, 200, 100);
            //Set the color and fill style of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.SeaGreen;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Rotate the shape to 90 degree
            shape.Rotation = 90;

            //Save the document to Pptx file
            string result = "SetRadiusOfRoundedRectangle.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}