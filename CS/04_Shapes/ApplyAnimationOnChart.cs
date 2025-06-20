using System;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing.Animation;

namespace ApplyAnimationOnChart
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
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\AnimationChart.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            //Get chart
            IShape shape = slide.Shapes[0];
            if (shape is IChart)
            { 
                //Apply Wipe animation effect to the chart
                AnimationEffect effect = slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Wipe);
                //Set the BuildType as Series
                effect.GraphicAnimation.BuildType = GraphicBuildType.BuildAsSeries;
            }
           
            //Save the document
            String result = "ApplyAnimationOnChart.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the document
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