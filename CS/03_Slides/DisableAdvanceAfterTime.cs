using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace DisableAdvanceAfterTime
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a Presentation object
            Presentation ppt = new Presentation();

            // Load the PPT file from the specified path
            ppt.LoadFromFile(@"..\..\..\..\..\..\..\Data\DisableAdvanceAfterTime.pptx");

            // Get the first slide and disable the selected advance after time setting
            ppt.Slides[0].SlideShowTransition.SelectedAdvanceAfterTime = false;

            // Specify the name for the output PowerPoint presentation file.
            string result = "output.pptx";

            // Save the modified PPT to the specified path
            ppt.SaveToFile(result, FileFormat.Pptx2013);

            // Dispose of the Presentation object to free up resources
            ppt.Dispose();

            // Launch the saved file
            PresentationDocViewer(result);
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