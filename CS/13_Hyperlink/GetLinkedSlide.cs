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

namespace GetLinkedSlide
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

            //Load ppt file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\linkedSlide.pptx");

            //Get the second slide
            ISlide slide = presentation.Slides[1];

            //Get the first shape of the second slide
            IAutoShape shape = slide.Shapes[0] as IAutoShape;

            //Get the linked slide index
            if (shape.Click.ActionType == HyperlinkActionType.GotoSlide)
            {
                ISlide targetSlide = shape.Click.TargetSlide;
                MessageBox.Show("Linked slide number = " + targetSlide.SlideNumber);
            }
        }
    }
}