using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace HeaderAndFooter
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

            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\HeaderAndFooter.pptx");

            //Add footer
            presentation.SetFooterText("Demo of Spire.Presentation");
           
            //Set the footer visible
            presentation.FooterVisible = true;

            //Set the page number visible
            presentation.SlideNumberVisible = true;

            //Set the date visible
            presentation.DateTimeVisible = true;
        
            //Save the document
            presentation.SaveToFile("HeaderAndFooter_result.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("HeaderAndFooter_result.pptx");
        }
    }
}