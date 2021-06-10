using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace GetSectionIndex
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
            Presentation ppt = new Presentation();

            //Load the document from disk
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\AddSection.pptx");
            
            Section section = ppt.SectionList[0];
            
            //Get the index of the section
            int index = ppt.SectionList.IndexOf(section);
            MessageBox.Show("The section index is: " + index);
        }
	
    }
}