using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;

namespace AccessSmartArtLayout
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the PPT
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.pptx");

            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ISmartArt)
                {
                    //Get the SmartArt and collect nodes
                    ISmartArt sa = shape as ISmartArt;
                    //Check SmartArt Layout
                    String layout = sa.LayoutType.ToString();
                    MessageBox.Show("SmartArt layout type is "+layout);
                }
            }     
        }
    }
}