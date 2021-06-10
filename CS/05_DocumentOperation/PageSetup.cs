using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace PageSetup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Set the size of slides
            presentation.SlideSize.Size = new SizeF(600,600);
            presentation.SlideSize.Orientation = SlideOrienation.Portrait;
            presentation.SlideSize.Type = SlideSizeType.Custom;

            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Append new shape
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 150, 400, 200);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

            //Add text to shape
            shape.AppendTextFrame("The sample demonstrates how to set slide size.");

            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Myriad Pro");
            shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(36,64,97);

            //Save the document
            presentation.SaveToFile("PageSetup.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("PageSetup.pptx");
        }
    }
}