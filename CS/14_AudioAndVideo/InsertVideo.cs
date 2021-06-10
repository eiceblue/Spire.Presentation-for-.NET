using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using System.IO;
using Spire.Presentation;

namespace InsertVideo
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\InsertVideo.pptx");

            //Add title
            RectangleF rec_title = new RectangleF(50, 280, 160, 50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.ShapeStyle.LineColor.Color = Color.Transparent;

            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "Video:";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 32;
            para_title.TextRanges[0].IsBold = TriState.True;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(68, 68, 68);
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //Insert video into the document
            RectangleF videoRect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 125, 240, 150, 150);
            IVideo video = presentation.Slides[0].Shapes.AppendVideoMedia(Path.GetFullPath(@"..\..\..\..\..\..\Data\Video.mp4"), videoRect);
            video.PictureFill.Picture.Url = @"..\..\..\..\..\..\Data\Video.png";
            
            //Save the document
            presentation.SaveToFile("video.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("video.pptx");
        }
    }
}