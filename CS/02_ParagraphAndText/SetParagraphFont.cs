using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;
using Spire.Presentation.Drawing;

namespace SetParagraphFont
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

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Az2.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Access the first and second placeholder in the slide and typecasting it as AutoShape
            ITextFrameProperties tf1 = ((IAutoShape)slide.Shapes[0]).TextFrame;
            ITextFrameProperties tf2 = ((IAutoShape)slide.Shapes[1]).TextFrame;

            // Access the first Paragraph
            TextParagraph para1 = tf1.Paragraphs[0];
            TextParagraph para2 = tf2.Paragraphs[0];

            //Justify the paragraph
            para2.Alignment = TextAlignmentType.Justify;

            //Access the first text range
            TextRange  textRange1 = para1.FirstTextRange;
            TextRange textRange2 = para2.FirstTextRange;

            //Define new fonts
            TextFont fd1 = new TextFont("Elephant");
            TextFont fd2 = new TextFont("Castellar");
 
            // Assign new fonts to text range
            textRange1.LatinFont = fd1;
            textRange2.LatinFont = fd2;

            // Set font to Bold
            textRange1.Format.IsBold = TriState.True;
            textRange2.Format.IsBold = TriState.False;

            // Set font to Italic
            textRange1.Format.IsItalic = TriState.False;
            textRange2.Format.IsItalic = TriState.True;

            // Set font color
            textRange1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange1.Fill.SolidColor.Color = Color.Purple;
            textRange2.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange2.Fill.SolidColor.Color = Color.Peru;

            string result = "SetParagraphFont_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);
            Viewer(result);
        }

        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}