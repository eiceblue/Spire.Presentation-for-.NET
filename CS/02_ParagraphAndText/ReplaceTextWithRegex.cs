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
using System.Text.RegularExpressions;

namespace ReplaceTextWithRegex
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

            //Load file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SomePresentation.pptx");

            //Regex for all words
            Regex regex = new Regex(@"\d+.\d+|\w+");

            //New string value
            string newvalue = "This is the test!";

            //Loop and replace
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    shape.ReplaceTextWithRegex(regex, newvalue);
                }
            }
            //Save the file
            String result = "ReplaceTextWithRegex_result.pptx";
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