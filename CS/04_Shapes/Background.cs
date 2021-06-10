using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace Background
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

            //Set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\backgroundImg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);

            //Add title
            RectangleF rec_title = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 70, 380, 50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.Line.FillType = FillFormatType.None;
            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "Background Sample";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Lucida Sans Unicode");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.DarkSlateBlue ;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //Add new shape to PPT document
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 155, 600, 200);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            shape.Line.FillType = FillFormatType.None;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

            TextParagraph para = new TextParagraph();
            para.Text = "Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc.";
            
            para.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para.TextRanges[0].Fill.SolidColor.Color = Color.CadetBlue;
            para.TextRanges[0].FontHeight = 26;
            shape.TextFrame.Paragraphs.Append(para);

            //Save the document
            presentation.SaveToFile("Background_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("Background_result.pptx");
        }
    }
}