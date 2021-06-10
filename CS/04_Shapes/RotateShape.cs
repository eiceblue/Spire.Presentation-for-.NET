using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RotateShape
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\RotateShape.pptx");

            //Get the shapes 
            IAutoShape shape = ppt.Slides[0].Shapes[0] as IAutoShape;

            //Set the rotation
            shape.Rotation =60;

            (ppt.Slides[0].Shapes[1] as IAutoShape).Rotation = 120;
            (ppt.Slides[0].Shapes[2] as IAutoShape).Rotation = 180;
            (ppt.Slides[0].Shapes[3] as IAutoShape).Rotation = 240;

            //Save the document
            ppt.SaveToFile("RotateShape_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RotateShape_result.pptx");
        }

    }
}
