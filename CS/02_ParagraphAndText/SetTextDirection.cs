using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetTextDirection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();

            //Append a shape with text to the first slide
            IAutoShape textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 70, 100, 400));
            textboxShape.ShapeStyle.LineColor.Color = Color.Transparent;
            textboxShape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textboxShape.Fill.SolidColor.Color = Color.LightBlue;
            textboxShape.TextFrame.Text = "You Are Welcome Here";
            //Set the text direction to vertical
            textboxShape.TextFrame.VerticalTextType = VerticalTextType.Vertical;

            //Append another shape with text to the slide
            textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(350, 70, 100, 400));
            textboxShape.ShapeStyle.LineColor.Color = Color.Transparent;
            textboxShape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textboxShape.Fill.SolidColor.Color = Color.LightGray;
            //Append some asian characters
            textboxShape.TextFrame.Text = "ª∂”≠π‚¡Ÿ";
            //Set the VerticalTextType as EastAsianVertical to aviod rotating text 90 degrees
            textboxShape.TextFrame.VerticalTextType = VerticalTextType.EastAsianVertical;

            //Save the document
            string result = "SetTextDirection.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
		}
	
		private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}