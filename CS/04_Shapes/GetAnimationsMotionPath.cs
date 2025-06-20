using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Animation;

namespace GetAnimationsMotionPath
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document
            Presentation presentation = new Presentation();
            //Load the file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\GetAnimationsMotionPath.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            //Get the first shape
            IShape shape = slide.Shapes[0];
            //Create a StringBuilder to save the tracks
            StringBuilder sb = new StringBuilder();
            int i = 1;
            //Traverse all animations
            foreach (AnimationEffect effect in shape.Slide.Timeline.MainSequence)
            {
                if (effect.ShapeTarget.Equals(shape as Shape))
                {
                    //Get MotionPath
                    MotionPath path = ((AnimationMotion)effect.CommonBehaviorCollection[0]).Path;
                    //Get all points in the path
                    foreach (MotionCmdPath motionCmdPath in path)
                    {
                        PointF[] points = motionCmdPath.Points;
                        MotionCommandPathType type = motionCmdPath.CommandType;
                        if (points != null)
                        {
                            foreach (PointF point in points)
                            {
                                sb.AppendLine(i+"  MotionType: " + type + " -> X: " + point.X + ", Y: " + point.Y);
                            }
                            i++;
                        }
                    }
                }
            }
            string result = "GetAnimationsMotionPath.txt";
            File.WriteAllText(result, sb.ToString());
            System.Diagnostics.Process.Start(result);
        }
    }
}