using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetRightToLeftColumns
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
            //Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\TwoColumns.pptx");

            //Get the second shape
            IAutoShape shape = ppt.Slides[0].Shapes[1] as IAutoShape;
            //Set columns style to right-to-left
            shape.TextFrame.RightToLeftColumns = true;

            //Save the document
            string result = "SetRightToLeftColumns.pptx";
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