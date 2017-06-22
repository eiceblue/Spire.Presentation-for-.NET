using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddWatermark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create a PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\table.pptx");

            //get the size of watermark string
            Graphics gc = this.CreateGraphics();
            SizeF size = gc.MeasureString("E-iceblue", new Font("Arial", 45));
            
            //define a rectangle range
            RectangleF rect = new RectangleF((presentation.SlideSize.Size.Width - size.Width) / 2, (presentation.SlideSize.Size.Height - size.Height) / 2, size.Width, size.Height);
            
            //add a rectangle shape with a defined range
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, rect);
            
            //set the style of shape
            shape.Fill.FillType = FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Rotation = -45;
            shape.Locking.SelectionProtection = true;
            shape.Line.FillType = FillFormatType.None;
            
            //add text to shape
            shape.TextFrame.Text = "E-iceblue";
            TextRange textRange = shape.TextFrame.TextRange;
            //set the style of the text range
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.FromArgb(120, Color.Black);
            textRange.FontHeight = 45;

            presentation.SaveToFile("Watermark.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("Watermark.pptx");
        }
    }
}
