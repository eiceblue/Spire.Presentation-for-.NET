using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Collections;
using Spire.Presentation.Drawing.Animation;

namespace SetAnimationRepeatType
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\Animation.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            AnimationEffectCollection animations = slide.Timeline.MainSequence;

            animations[0].Timing.AnimationRepeatType = AnimationRepeatType.UtilEndOfSlide;

            String result = "SetAnimationRepeatType_result.pptx";

            //Save to file.
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