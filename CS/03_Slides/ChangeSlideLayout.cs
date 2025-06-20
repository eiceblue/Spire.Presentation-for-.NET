using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ChangeSlideLayout
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
			Presentation presentation=new Presentation();

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\ChangeSlideLayout.pptx");

			//Change the layout of slide
			presentation.Slides[1].Layout = presentation.Masters[0].Layouts[4];

            //Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Output.pptx");

        }
    }
}
