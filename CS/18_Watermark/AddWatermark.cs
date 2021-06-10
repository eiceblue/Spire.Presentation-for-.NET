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
            //Create a PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddWatermark.pptx");

            //Get the size of the watermark string
            Graphics gc = this.CreateGraphics();
            SizeF size = gc.MeasureString("E-iceblue", new Font("Lucida Sans Unicode", 50));
            
            //Define a rectangle range
            RectangleF rect = new RectangleF((presentation.SlideSize.Size.Width - size.Width) / 2, (presentation.SlideSize.Size.Height - size.Height) / 2, size.Width, size.Height);
            
            //Add a rectangle shape with a defined range
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, rect);
            
            //Set the style of the shape
            shape.Fill.FillType = FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Rotation = -45;
            shape.Locking.SelectionProtection = true;
            shape.Line.FillType = FillFormatType.None;
            
            //Add text to the shape
            shape.TextFrame.Text = "E-iceblue";
            TextRange textRange = shape.TextFrame.TextRange;
            //Set the style of the text range
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.FromArgb(120, Color.HotPink);
            textRange.FontHeight = 50;
            
			//Save the document and launch
            presentation.SaveToFile("Watermark_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("Watermark_result.pptx");
        }
    }
}
