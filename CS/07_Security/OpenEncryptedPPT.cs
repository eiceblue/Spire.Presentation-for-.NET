using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OpenEncryptedPPT
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
            
            //Load the PPT with password
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\OpenEncryptedPPT.pptx", FileFormat.Pptx2010, textBox1.Text);

            //Save as a new PPT with original password
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Output.pptx");

        }
    }
}
