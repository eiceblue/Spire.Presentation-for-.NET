using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing.Animation;
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
            para_title.Text = "Bullets";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //append new shape
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(110, 155, 400, 280));
            shape.Fill.FillType = FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.TextFrame.Paragraphs.RemoveAt(0);

            string[] str = new string[] { "Spire.Office for .NET", "Spire.Doc for .NET", "Spire.XLS for .NET", "Spire.PDF for .NET", "Spire.Presentation for .NET", "Spire.Barcode for .NET", "Spire.DataExport for .NET", "Spire.DocViewer for .NET", "Spire.PDFViewer for .NET" };
            foreach (string txt in str)
            {
                TextParagraph textParagraph = new TextParagraph();
                textParagraph.Text = txt;
                textParagraph.Alignment = TextAlignmentType.Left;
                textParagraph.Indent = 35;

                //set the Bullets
                textParagraph.BulletType = TextBulletType.Numbered;
                textParagraph.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod;
                shape.TextFrame.Paragraphs.Append(textParagraph);
            }

            //set the font and fill style
            foreach (TextParagraph paragraph in shape.TextFrame.Paragraphs)
            {
                paragraph.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
                paragraph.TextRanges[0].FontHeight = 24;
                paragraph.TextRanges[0].Fill.FillType = FillFormatType.Solid;
                paragraph.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            }

            //save the document
            presentation.SaveToFile("bullets.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("bullets.pptx");
        }
    }
}