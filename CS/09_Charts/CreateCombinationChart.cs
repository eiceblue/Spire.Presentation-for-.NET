using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateCombinationChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a presentation instance
            Presentation presentation = new Presentation();

			//Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect2 = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;
			
            //Insert a column clustered chart
            RectangleF rect = new RectangleF(100, 100, 550, 320);
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect);

            //Set chart title
            chart.ChartTitle.TextProperties.Text = "Monthly Sales Report";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            //Create a datatable
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add(new DataColumn("Month", Type.GetType("System.String")));
            dataTable.Columns.Add(new DataColumn("Sales", Type.GetType("System.Int32")));
            dataTable.Columns.Add(new DataColumn("Growth rate", Type.GetType("System.Decimal")));
            dataTable.Rows.Add("January", 200, 0.6);
            dataTable.Rows.Add("February", 250, 0.8);
            dataTable.Rows.Add("March", 300, 0.6);
            dataTable.Rows.Add("April", 150, 0.2);
            dataTable.Rows.Add("May", 200, 0.5);
            dataTable.Rows.Add("June", 400, 0.9);

            //Import data from datatable to chart data
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                chart.ChartData[0, c].Text = dataTable.Columns[c].Caption;
            }
            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                object[] datas = dataTable.Rows[r].ItemArray;
                for (int c = 0; c < datas.Length; c++)
                {
                    chart.ChartData[r + 1, c].Value = datas[c];

                }
            }

            //Set series labels
            chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];

            //Set categories labels    
            chart.Categories.CategoryLabels = chart.ChartData["A2", "A7"];

            //Assign data to series values
            chart.Series[0].Values = chart.ChartData["B2", "B7"];
            chart.Series[1].Values = chart.ChartData["C2", "C7"];

            //Change the chart type of serie 2 to line with markers
            chart.Series[1].Type = ChartType.LineMarkers;

            //Plot data of series 2 on the secondary axis
            chart.Series[1].UseSecondAxis = true;

            //Set the number format as percentage 
            chart.SecondaryValueAxis.NumberFormat = "0%";

            //Hide gridlinkes of secondary axis
            chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;

            //Set overlap
            chart.OverLap = -50;

            //Set gapwidth
            chart.GapWidth = 200;

            //Save to file
            presentation.SaveToFile("CombinationChart_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("CombinationChart_result.pptx");
        }
    }
}
