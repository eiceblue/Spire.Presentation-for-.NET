using Spire.Presentation;
using Spire.Presentation.Collections;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AppendSlideWithMasterLayout
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

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AppendSlideWithMasterLayout.pptx");

            //Get the master
            IMasterSlide master = presentation.Masters[0];

            //Get master layout slides
            IMasterLayouts masterLayouts = master.Layouts;
            ActiveSlide layoutSlide = masterLayouts[1] as ActiveSlide;

            //Append a rectangle to the layout slide
            IAutoShape shape = layoutSlide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(10, 50, 100, 80));

            //Add a text into the shape and set the style
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.AppendTextFrame("Layout slide 1");
            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Black");
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.CadetBlue;

            //Append new slide with master layout
            presentation.Slides.Append(presentation.Slides[0], master.Layouts[1]);

            //Another way to append new slide with master layout
            presentation.Slides.Insert(2, presentation.Slides[1], master.Layouts[1]);

            //Save the document
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("output.pptx");
        }
    }
}
