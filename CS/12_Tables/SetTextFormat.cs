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

namespace SetTextFormat
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Table.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            StringBuilder str = new StringBuilder();
            foreach (IShape shape in slide.Shapes)
            {
                //Verify if it is table
                if (shape is ITable)
                {
                    ITable table = (ITable)shape;

                    Cell cell1 = table.TableRows[0][0];
                    //Set table cell's text alignment type 
                    cell1.TextAnchorType = TextAnchorType.Top;
                    //Set italic style
                    cell1.TextFrame.TextRange.Format.IsItalic = TriState.True;

                    Cell cell2 = table.TableRows[1][0];
                    //Set table cell's foreground color
                    cell2.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    cell2.TextFrame.TextRange.Fill.SolidColor.Color = Color.Green;
                    //Set table cell's background color
                    cell2.FillFormat.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    cell2.FillFormat.SolidColor.Color = Color.LightGray;
                   

                    Cell cell3 = table.TableRows[2][2];
                    //Set table cell's font and font size
                    cell3.TextFrame.TextRange.FontHeight = 12;
                    cell3.TextFrame.TextRange.LatinFont = new TextFont("Arial Black");
                    cell3.TextFrame.TextRange.HighlightColor.Color = Color.YellowGreen;
                  

                    Cell cell4 = table.TableRows[2][1];
                    //Set table cell's margin and borders
                    cell4.MarginLeft = 20;
                    cell4.MarginTop = 30;
                    cell4.BorderTop.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    cell4.BorderTop.SolidFillColor.Color = Color.Red;
                    cell4.BorderBottom.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    cell4.BorderBottom.SolidFillColor.Color = Color.Red;
                    cell4.BorderLeft.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    cell4.BorderLeft.SolidFillColor.Color = Color.Red;
                    cell4.BorderRight.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    cell4.BorderRight.SolidFillColor.Color = Color.Red;            
                }  
            }

            string result = "SetTextFormat_result.pptx";
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