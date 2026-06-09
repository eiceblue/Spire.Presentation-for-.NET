using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Spire.Presentation;

namespace HighlightTextByRegex
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SomePresentation.pptx");

            // Create a regex pattern to match the word 
            Regex regex = new Regex(@"\bSpire\b");

            //Get the specified shape
            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];

            // Initialize text highlighting options
            TextHighLightingOptions options = new TextHighLightingOptions();
            shape.TextFrame.HighLightRegex(regex, Color.Red, options);

            //Save to file
            String result = "HighlightTextByRegex_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PPT file
            System.Diagnostics.Process.Start(result);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}