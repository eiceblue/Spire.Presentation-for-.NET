using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace LinkToASpecificSlide
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Append a slide to it.
            presentation.Slides.Append();

            //Add a shape to the second slide.
            IAutoShape shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(10, 50, 200, 50));
            shape.Fill.FillType = FillFormatType.None;
            shape.Line.FillType = FillFormatType.None;
            shape.TextFrame.Text = "Jump to the first slide";

            //Create a hyperlink based on the shape and the text on it, linking to the first slide.
            ClickHyperlink hyperlink = new ClickHyperlink(presentation.Slides[0]);
            shape.Click = hyperlink;
            shape.TextFrame.TextRange.ClickAction = hyperlink;

            String result = "Result-LinkToASpecificSlide.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}