using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace HyperlinkOutlineStyle
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

            //Add new shape to PPT document
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 255, 120, 400, 100);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            shape.Fill.FillType =FillFormatType.None;
            shape.Line.FillType = FillFormatType.None;

            //Add a paragraph with hyperlink
            TextParagraph para1 = new TextParagraph();
            TextRange tr1 = new TextRange("Click to know more about Spire.Presentation");
            tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
            para1.TextRanges.Append(tr1);

            //Set the format of textrange
            tr1.Format.FontHeight = 20f;
            tr1.IsItalic = TriState.True;
         
            //Set the outline format of textrange
            tr1.TextLineFormat.FillFormat.FillType = FillFormatType.Solid;
            tr1.TextLineFormat.FillFormat.SolidFillColor.Color = Color.LightSeaGreen;
            tr1.TextLineFormat.JoinStyle = LineJoinType.Round;
            tr1.TextLineFormat.Width = 2f;
            
            //Add the paragraph to shape
            shape.TextFrame.Paragraphs.Append(para1); 
            shape.TextFrame.Paragraphs.Append(new TextParagraph());

            //Save the document
            string result = "HyperlinkOutlineStyle.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}