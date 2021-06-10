using Spire.Presentation.Collections;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ChangeTextStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeTextStyle.pptx");

            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];
            ParagraphCollection paras = shape.TextFrame.Paragraphs;

            //Set the style for the text content in the first paragraph
            foreach (TextRange tr in paras[0].TextRanges)
            {
                tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                tr.Fill.SolidColor.Color = Color.ForestGreen;
                tr.LatinFont = new TextFont("Lucida Sans Unicode");
                tr.FontHeight = 14;
            }

            //Set the style for the text content in the third paragraph
            foreach (TextRange tr in paras[2].TextRanges)
            {
                tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                tr.Fill.SolidColor.Color = Color.CornflowerBlue;
                tr.LatinFont = new TextFont("Calibri");
                tr.FontHeight = 16;
                tr.TextUnderlineType = TextUnderlineType.Dashed;
            }
            
            //Save the document
            presentation.SaveToFile("ChangeTextStyle_result.pptx", FileFormat.Pptx2007);
            System.Diagnostics.Process.Start("ChangeTextStyle_result.pptx");
        }
    }
}