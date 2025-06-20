using System;

using System.Windows.Forms;
using Spire.Presentation;

namespace ShowMasterBackgroundGraphics
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {            
            // Create a Presentation object and load the input file 
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ShowMasterBackgroundGraphics.pptx");

            // Set whether to show the background graphics of the slide master
            presentation.Slides[0].Layout.ShowMasterShapes = true;

            // Save file 
            presentation.SaveToFile("ShowMasterBackgroundGrapics_output.pptx",FileFormat.Pptx2019);

            //Dispose
            presentation.Dispose();

            System.Diagnostics.Process.Start("ShowMasterBackgroundGrapics_output.pptx");
        }
    }
}