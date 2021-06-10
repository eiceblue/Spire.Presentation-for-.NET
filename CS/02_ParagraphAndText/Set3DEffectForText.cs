using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace Set3DEffectForText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a new presentation object
            Presentation ppt = new Presentation();

            //Get the first slide
            ISlide slide = ppt.Slides[0];

            //Append a new shape to slide and set the line color and fill type
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(30, 40, 650, 200));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = FillFormatType.None;

            //Add text to the shape
            shape.AppendTextFrame("This demo shows how to add 3D effect text to Presentation slide");

            //Set the color of text in shape
            TextRange textRange = shape.TextFrame.TextRange;
            textRange.Fill.FillType = FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.LightBlue;

            //Set the Font of text in shape
            textRange.FontHeight = 40;
            textRange.LatinFont = new TextFont("Gulim");

            //Set 3D effect for text
            shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = PresetMaterialType.Matte;
            shape.TextFrame.TextThreeD.LightRig.PresetType = PresetLightRigType.Sunrise;
            shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = BevelPresetType.Circle;
            shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = Color.Green;
            shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3;

            //Save the document
            string result = "Set3DEffectForText.pptx";
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