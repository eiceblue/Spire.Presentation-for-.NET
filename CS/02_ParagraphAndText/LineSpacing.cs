using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;
using Spire.Presentation.Drawing;

namespace LineSpacing
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

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Az.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            //Add a shape 
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, presentation.SlideSize.Size.Width-100,300));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.TextFrame.Paragraphs.Clear();

            //Add text
            shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers to"
            +"create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform."
            +"From Spire.Presentation v 3.7.5, Spire.Presentation starts to support .NET Core, .NET standard.");
            //Set font and color of text
            TextRange textRange = shape.TextFrame.TextRange;
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.BlueViolet;
            textRange.FontHeight =20;
            textRange.LatinFont = new TextFont("Lucida Sans Unicode");
           
            //Set properties of paragraph
            shape.TextFrame.Paragraphs[0].SpaceBefore = 100;
            shape.TextFrame.Paragraphs[0].SpaceAfter = 100;
            shape.TextFrame.Paragraphs[0].LineSpacing = 150;       
     
            string result = "LineSpacing_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);
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