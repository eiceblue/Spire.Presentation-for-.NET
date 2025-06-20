using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetRadiusForRoundedRectangle
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

            //Get the first slide 
            ISlide slide = presentation.Slides[0];

            //Insert a rectangle with four round corners and set its radius
            IAutoShape shape1 = slide.Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 50, 150, 150));
            shape1.SetRoundRadius(shape1.Width / 3);

            //Insert a rectangle with one round corner and set its radius
            IAutoShape shape2 = slide.Shapes.AppendShape(ShapeType.OneRoundCornerRectangle, new RectangleF(250, 50, 150, 150));
            shape2.SetRoundRadius(shape2.Width / 3);

            //Insert a rectangle with one round corner and which one round cornet is snipped and set its radius
            IAutoShape shape3 = slide.Shapes.AppendShape(ShapeType.OneSnipOneRoundCornerRectangle, new RectangleF(450, 50, 150, 150));
            shape3.SetRoundRadius(shape3.Width / 3);

            //Insert a rectangle with two diagonal round corners and set its radius
            IAutoShape shape4 = slide.Shapes.AppendShape(ShapeType.TwoDiagonalRoundCornerRectangle, new RectangleF(50, 250, 150, 150));
            shape4.SetRoundRadius(shape4.Width / 3);

            //Insert a rectangle with two same side round corners and set its radius
            IAutoShape shape5 = slide.Shapes.AppendShape(ShapeType.TwoSamesideRoundCornerRectangle, new RectangleF(250, 250, 150, 150));
            shape5.SetRoundRadius(shape5.Width / 3);


            //Save to file.
            String result = "output.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}