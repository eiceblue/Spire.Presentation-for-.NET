using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Collections;
using Spire.Presentation.Drawing.Animation;

namespace CustomPathAnimation
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
            Presentation ppt = new Presentation();

            //Add shape
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(0, 0, 200, 200));

            //Add animation
            AnimationEffect effect = ppt.Slides[0].Timeline.MainSequence.
                AddEffect(shape, AnimationEffectType.PathUser);
            CommonBehaviorCollection common = effect.CommonBehaviorCollection;
            AnimationMotion motion = (AnimationMotion)common[0];
            motion.Origin = AnimationMotionOrigin.Layout;
            motion.PathEditMode = AnimationMotionPathEditMode.Relative;

            //Add moin path
            MotionPath moinPath = new MotionPath();
            moinPath.Add(MotionCommandPathType.MoveTo, new PointF[] { new PointF(0, 0) }, MotionPathPointsType.CurveAuto, true);
            moinPath.Add(MotionCommandPathType.LineTo, new PointF[] { new PointF(0.1f, 0.1f) }, MotionPathPointsType.CurveAuto, true);
            moinPath.Add(MotionCommandPathType.LineTo, new PointF[] { new PointF(-0.1f, 0.2f) }, MotionPathPointsType.CurveAuto, true);
            moinPath.Add(MotionCommandPathType.End, new PointF[] { }, MotionPathPointsType.CurveStraight, true);
            motion.Path = moinPath;

            //Save the document
            string outputFile = "result.pptx";
            ppt.SaveToFile(outputFile, FileFormat.Pptx2010);
            ppt.Dispose();

            //Launch the PPT file
            FileViewer(outputFile);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
