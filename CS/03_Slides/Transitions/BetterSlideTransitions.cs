using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
namespace BetterSlideTransitions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the PPT
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\SetTransitions.pptx");

            //Set the first slide transition as circle
            presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle;

            // Set the transition time of 3 seconds
            presentation.Slides[0].SlideShowTransition.AdvanceOnClick = true;
            presentation.Slides[0].SlideShowTransition.AdvanceAfterTime = 3000;

            //Set the second slide transition as comb and set the speed 
            presentation.Slides[1].SlideShowTransition.Type = TransitionType.Comb;
            presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow;

            // Set the transition time of 5 seconds
            presentation.Slides[1].SlideShowTransition.AdvanceOnClick = true;
            presentation.Slides[1].SlideShowTransition.AdvanceAfterTime = 5000;

            // Set the third slide transition as zoom
            presentation.Slides[2].SlideShowTransition.Type = TransitionType.Zoom;

            // Set the transition time of 7 seconds
            presentation.Slides[2].SlideShowTransition.AdvanceOnClick = true;
            presentation.Slides[2].SlideShowTransition.AdvanceAfterTime = 7000;

          
            String result = "BetterSlideTransitions_result.pptx";
            //Save the file
            presentation.SaveToFile(result, FileFormat.Pptx2010);

            Viewer(result);
        }

        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}