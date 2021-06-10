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

namespace InsertAudio
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\InsertAudio.pptx");
           
            //Add title
            RectangleF rec_title = new RectangleF(50, 240, 160,50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.ShapeStyle.LineColor.Color = Color.Transparent;
            
            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "Audio:";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 32;
            para_title.TextRanges[0].IsBold = TriState.True;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(68,68,68);
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //Insert audio into the document
            RectangleF audioRect = new RectangleF(220, 240, 80, 80);
            presentation.Slides[0].Shapes.AppendAudioMedia(Path.GetFullPath(@"..\..\..\..\..\..\Data\Music.wav"), audioRect);

            //Save the document
            presentation.SaveToFile("Audio.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Audio.pptx");
        }
    }
}