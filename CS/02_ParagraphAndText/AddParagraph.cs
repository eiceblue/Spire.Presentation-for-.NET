using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace AddParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();

            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            ppt.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Append a new shape
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 70, 620, 150));
            shape.Fill.FillType = FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Set the alignment of paragraph
            shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            //Set the indent of paragraph
            shape.TextFrame.Paragraphs[0].Indent = 50;
            //Set the linespacing of paragraph
            shape.TextFrame.Paragraphs[0].LineSpacing = 150;
            //Set the text of paragraph
            shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue.";

            //Set the Font
            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Rounded MT Bold");
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;

            //Save and launch the document
            string result = "AddParagraph.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
	}
	
	private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}