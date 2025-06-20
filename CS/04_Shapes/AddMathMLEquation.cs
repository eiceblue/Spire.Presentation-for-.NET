using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddMathMLEquation
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
            Presentation ppt = new Presentation();

            //Set the mathML code
            String mathMLCode = "<mml:math xmlns:mml=\"http://www.w3.org/1998/Math/MathML\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">" + "<mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:msqrt><mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:msqrt><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:math>";

            //Add a shape
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new Rectangle(30, 100, 400, 30));
            shape.TextFrame.Paragraphs.Clear();

            //Add the mathml equation paragraph
            TextParagraph tp = shape.TextFrame.Paragraphs.AddParagraphFromMathMLCode(mathMLCode);

            //Save the document
            String outputFile = "result.pptx";
            ppt.SaveToFile(outputFile, FileFormat.Pptx2013);
            ppt.Dispose();
            PresentationDocViewer(outputFile);
        }
        private static void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}