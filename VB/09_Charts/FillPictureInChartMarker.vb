Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing

Namespace FillPictureInChartMarker
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ChartSample4.pptx")

			'Get chart on the first slide
			Dim Chart As IChart = TryCast(ppt.Slides(0).Shapes(0), IChart)

			'Load image file in ppt
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Logo.png")
			Dim IImage As IImageData = ppt.Images.Append(image)

			'Create a ChartDataPoint object and specify the index
			Dim dataPoint As New ChartDataPoint(Chart.Series(0))
			dataPoint.Index = 0

			'Fill picture in marker
			dataPoint.MarkerFill.Fill.FillType = FillFormatType.Picture
			dataPoint.MarkerFill.Fill.PictureFill.Picture.EmbedImage = IImage

			'Set marker size
			dataPoint.MarkerSize = 20

			'Add the data point in series
			Chart.Series(0).DataPoints.Add(dataPoint)

			Dim result As String = "FillPictureInChartMarker_result.pptx"
			'Save the document
			ppt.SaveToFile(result, FileFormat.Pptx2010)

			'Launch the result file
			PPTDocViewer(result)
		End Sub

		Private Sub PPTDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace