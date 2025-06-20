using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;
using Spire.Presentation.Charts;

namespace AddAndDetectMathEquations
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create Presentation
            Presentation presentation = new Presentation();

            //Math code
            string latexMathCode = @"x^{2}+\sqrt{x^{2}+1}=2";

            //Append a shape
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, new RectangleF(30, 100, 400, 30));
            shape.TextFrame.Paragraphs.Clear();

            //Add math equation
            TextParagraph tp = shape.TextFrame.Paragraphs.AddParagraphFromLatexMathCode(latexMathCode);

            //Detect if the slide contains math equation
            for (int i = 0; i < presentation.Slides[0].Shapes.Count; i++)
            {

                if (presentation.Slides[0].Shapes[i] is IAutoShape)
                {
                    bool containMathEquation = (presentation.Slides[0].Shapes[i] as IAutoShape).ContainMathEquation;
                    MessageBox.Show("The first slide contains math equations: " + containMathEquation);
                }
            }

            //Save the file
            String result = "AddAndDetectMathEquations_result.pptx";
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);

            //Launching the result file.
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