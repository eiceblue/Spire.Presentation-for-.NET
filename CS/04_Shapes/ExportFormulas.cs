using Spire.Presentation;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExportFormulas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Load the PPT file that contains formulas
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExportFormulas.pptx");

            // Initialize a StringBuilder to store all exported formulas
            StringBuilder stringBuilder = new StringBuilder();
            int count = 0;

            // Iterate through all slides and shapes to export every formula
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                ISlide slide = presentation.Slides[i];

                for (int j = 0; j < slide.Shapes.Count; j++)
                {
                    // Attempt to cast the shape to an AutoShape that can contain text
                    IAutoShape autoShape = slide.Shapes[j] as IAutoShape;

                    // Skip shapes that are not AutoShapes or have no text frame
                    if (autoShape == null || autoShape.TextFrame == null)
                    {
                        continue;
                    }

                    // Iterate through all paragraphs in the text frame
                    for (int k = 0; k < autoShape.TextFrame.Paragraphs.Count; k++)
                    {
                        TextParagraph paragraph = autoShape.TextFrame.Paragraphs[k];

                        try
                        {
                            // Export the formula from the paragraph in both MathML and LaTeX formats
                            string mathML = paragraph.ExportMathML();
                            string latex = paragraph.ExportLaTex();

                            // Skip if the paragraph does not contain any formula
                            if (string.IsNullOrEmpty(mathML) && string.IsNullOrEmpty(latex))
                            {
                                continue;
                            }

                            // Increment the formula counter and append the exported data
                            count++;
                            stringBuilder.AppendLine("Formula " + count + ":");
                            stringBuilder.AppendLine("  MathML:");
                            stringBuilder.AppendLine("  " + mathML);
                            stringBuilder.AppendLine("  LaTeX:");
                            stringBuilder.AppendLine("  " + latex);
                            stringBuilder.AppendLine();
                        }
                        catch
                        {
                            // Skip the paragraphs that contain no formula or have export errors
                            // Continue processing remaining paragraphs
                        }
                    }
                }
            }

            // Prepend a summary line indicating the total number of formulas exported
            stringBuilder.Insert(0, "Exported " + count + " formula(s) in total." + Environment.NewLine + Environment.NewLine);

            // Save all the exported formulas into a text file
            File.WriteAllText("ExportFormulas_out.txt", stringBuilder.ToString());

            // Launch the text file to display the exported formulas
            System.Diagnostics.Process.Start("ExportFormulas_out.txt");
        }

    }
}
