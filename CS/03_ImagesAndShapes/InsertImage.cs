using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;

namespace Spire.Presentation.Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //add title
            RectangleF rec_title = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 70, 400, 50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.ShapeStyle.LineColor.Color = Color.White;
            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "Insert Image";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //insert image to PPT
            string ImageFile2 = @"..\..\..\..\..\..\Data\PresentationNET.png";
            RectangleF rect1 = new RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 170, 110, 120);
            IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile2, rect1);
            image.Line.FillType = FillFormatType.None;

            //add new shape to PPT document
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 190, 155, 490, 160);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);

            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

            //add text to shape
            shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.");

            //set the font and fill style of text
            TextParagraph paragraph = shape.TextFrame.Paragraphs[0];
            paragraph.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            paragraph.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            paragraph.TextRanges[0].FontHeight = 20;
            paragraph.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
            paragraph.Alignment = TextAlignmentType.Left;

            //add new shape to PPT document
            RectangleF rec2 = new RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 320, 600, 130);
            IAutoShape shape2 = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec2);

            shape2.ShapeStyle.LineColor.Color = Color.White;
            shape2.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

            //add text to shape
            shape2.AppendTextFrame("Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc.");

            //set the font and fill style of text
            TextParagraph paragraph2 = shape2.TextFrame.Paragraphs[0];
            paragraph2.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            paragraph2.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            paragraph2.TextRanges[0].FontHeight = 20;
            paragraph2.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
            paragraph2.Alignment = TextAlignmentType.Left;

            //save the document
            presentation.SaveToFile("InsertPicture.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("InsertPicture.pptx");
        }
    }
}