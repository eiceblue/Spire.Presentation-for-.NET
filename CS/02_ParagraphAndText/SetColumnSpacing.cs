using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetColumnSpacing
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new PPT
            Presentation presentation = new Presentation();

            // Append a shape in the first slide
            ISlide slide = presentation.Slides[0];
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 70, 600, 400));
            shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            shape.Fill.FillType = FillFormatType.None;

            // Set column and column spacing
            shape.TextFrame.ColumnCount = 2;
            shape.TextFrame.ColumnSpacing = 20.50f;
            // Append text
            shape.TextFrame.Text = "\r\nSpire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers to create, read, write, modify, convert and Print PowerPoint documents on any .NET platform (Target .NET Framework, .NET Core, .NET Standard, .NET 5.0, .NET 6.0, Xamarin & Mono Android). As an independent PowerPoint .NET API, Spire.Presentation for .NET doesn't need Microsoft PowerPoint to be installed on machines.\r\n\r\n\r\nSpire.Presentation for .NET supports PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also supports exporting presentation slides to images (PNG, JPG, TIFF, EMF, SVG), PDF, XPS, HTML format etc.";
            foreach (TextParagraph paragraph in shape.TextFrame.Paragraphs)
            {
                foreach (TextRange textRange in paragraph.TextRanges)
                {
                    // Set font for text
                    textRange.Fill.FillType = FillFormatType.Solid;
                    textRange.Fill.SolidColor.Color = Color.Black;
                    textRange.FontHeight = 16;
                    textRange.LatinFont = new TextFont("Open Sans");
                }
            }

            String result = "Result-SetColumnSpacing.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);
            // Dispose the document
            presentation.Dispose();

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