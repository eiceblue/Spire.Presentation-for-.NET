using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetRectangleFormat
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

            //Add a shape
            RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 100, 100, 200, 100);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect);

            //Set the fill format of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.CadetBlue;

            //Set the fill format of line
            shape.Line.FillType = FillFormatType.Solid;
            shape.Line.SolidFillColor.Color = Color.DimGray;

            //Save the document
            string result = "SetRectangleFormat_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}