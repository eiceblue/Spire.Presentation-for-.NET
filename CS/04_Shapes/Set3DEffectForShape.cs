using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace Set3DEffectForShape
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

            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            ppt.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Add shape1 and fill it with color
            IAutoShape shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(150, 150, 150, 150));
            shape1.Fill.FillType = FillFormatType.Solid;
            shape1.Fill.SolidColor.KnownColor = KnownColors.SkyBlue;
            //Initialize a new instance of the 3-D class for shape1 and set its properties
            ShapeThreeD effect1 = shape1.ThreeD.ShapeThreeD;
            effect1.PresetMaterial = PresetMaterialType.Powder;
            effect1.TopBevel.PresetType = BevelPresetType.ArtDeco;
            effect1.TopBevel.Height = 4;
            effect1.TopBevel.Width = 12;
            effect1.BevelColorMode = BevelColorType.Contour;
            effect1.ContourColor.KnownColor = KnownColors.LightBlue;
            effect1.ContourWidth = 3.5;

            //Add shape2 and fill it with color
            IAutoShape shape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Pentagon, new RectangleF(400, 150, 150, 150));
            shape2.Fill.FillType = FillFormatType.Solid;
            shape2.Fill.SolidColor.KnownColor = KnownColors.LightGreen;
            //Initialize a new instance of the 3-D class for shape2 and set its properties
            ShapeThreeD effect2 = shape2.ThreeD.ShapeThreeD;
            effect2.PresetMaterial = PresetMaterialType.SoftEdge;
            effect2.TopBevel.PresetType = BevelPresetType.SoftRound;
            effect2.TopBevel.Height = 12;
            effect2.TopBevel.Width = 12;
            effect2.BevelColorMode = BevelColorType.Contour;
            effect2.ContourColor.KnownColor = KnownColors.LawnGreen;
            effect2.ContourWidth = 5;

            //Save the document
            string result = "Set3DEffectForShape.pptx";
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