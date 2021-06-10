Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace SetGradientBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\PPTSample_N.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Set the background to gradient
			slide.SlideBackground.Type = BackgroundType.Custom
			slide.SlideBackground.Fill.FillType = FillFormatType.Gradient

			'Add gradient stops
			slide.SlideBackground.Fill.Gradient.GradientStops.Append(0.1f, Color.LightSeaGreen)
			slide.SlideBackground.Fill.Gradient.GradientStops.Append(0.7f, Color.LightCyan)

			'Set gradient shape type
			slide.SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear

			'Set the angle
			slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45

			'Save the document
			Dim result As String = "SetGradientBackground_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace