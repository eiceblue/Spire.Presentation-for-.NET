using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace ReplaceAndFormatText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Presentation object.
            Presentation ppt = new Presentation();

            // Load a PowerPoint presentation from the specified file.
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\TextTemplate.pptx");

            // Create a new object to store the default text range formatting properties.
            DefaultTextRangeProperties format = new DefaultTextRangeProperties();

            // Set the IsBold property of the text range formatting to true, making the text bold.
            format.IsBold = TriState.True;

            // Set the FillType property of the text range fill to Solid, indicating a solid fill color.
            format.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;

            // Set the Color property of the solid fill color to red.
            format.Fill.SolidColor.Color = Color.Red;

            // Set the FontHeight property of the text range formatting to 25, indicating the font size.
            format.FontHeight = 25;

            // Replace all occurrences of the text "Spire.Presentation for .NET" with "Spire.PPT" and apply the specified formatting.
            ppt.ReplaceAndFormatText("Spire.Presentation for .NET", "Spire.PPT", format);

            // Specify the name for the output PowerPoint presentation file.
            string result = "output.pptx";

            // Save the modified presentation to the specified output file in the PPTX format compatible with PowerPoint 2016.
            ppt.SaveToFile(result, FileFormat.Pptx2016);

            // Dispose of the Presentation object to free up resources
            ppt.Dispose();

            // Launch the saved file
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