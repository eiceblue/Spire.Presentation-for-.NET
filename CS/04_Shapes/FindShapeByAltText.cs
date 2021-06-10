using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace FindShapeByAltText
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

            //Load document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\FindShapeByAltText.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Find shape in the slide
            IShape shape = FindShape(slide, "Shape1");

            if (shape != null)
            {
                MessageBox.Show(shape.Name);
            }

        }
        private IShape FindShape(ISlide slide, string altText)
        {
            //Loop through shapes in the slide
            foreach (IShape shape in slide.Shapes)
            {
                //Find the shape whose alternative text is altText
                if (shape.AlternativeText.CompareTo(altText) == 0)
                {
                    return shape;
                }
            }
            return null;
        }
    }
}