using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddSVGAsShapes
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

            // Add the SVG file as a shape onto the first slide of the presentation.
            presentation.Slides[0].Shapes.AddFromSVGAsShapes(@"..\..\..\..\..\..\Data\AddSVGAsShapes.svg");

            //Save the document
            string result = "AddSVGAsShapes.pptx";
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