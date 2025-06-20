using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Animation;

namespace GetAnimationEffectInfo
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

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\Animation.pptx");

            StringBuilder sb = new StringBuilder();
            //Travel each slide
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (AnimationEffect effect in slide.Timeline.MainSequence)
                {
                    //Get the animation effect type
                    AnimationEffectType animationEffectType = effect.AnimationEffectType;
                    sb.AppendLine("animation effect type:" + animationEffectType);

                    //Get the slide number where the animation is located
                    int slideNumber = slide.SlideNumber;
                    sb.AppendLine("slide number:" + slideNumber );

                    //Get the shape name
                    string shapeName = effect.ShapeTarget.Name;
                    sb.AppendLine("shape name:" + shapeName + "\n");
                                     
                }
            }

            //Save the information of animation effect
            String result = "AnimationEffectInfo.txt";
            File.WriteAllText(result, sb.ToString());

            System.Diagnostics.Process.Start(result);
        }
    }
}