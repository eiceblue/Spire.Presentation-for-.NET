Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetBrightnessAndTransparencyForGradientStop
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a Presentation object
			Dim presentation As New Presentation()

			' Append new shape - BentUpArrow
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.BentUpArrow, New RectangleF(470, 300, 150, 100))

			' Set the color of shape
			shape.Fill.FillType = FillFormatType.Gradient

			' Add gradient stops to create a gradient fill
			shape.Fill.Gradient.GradientStops.Append(0f, KnownColors.Olive)
			shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.PowderBlue)

			' Adjust the brightness and transparency of the first gradient stop
			shape.Fill.Gradient.GradientStops(0).Color.Brightness = 0.5f
			shape.Fill.Gradient.GradientStops(0).Color.Transparency = 0.5f

			' Set the line color of the shape
			shape.ShapeStyle.LineColor.Color = Color.White

			' Specify the name for the output PowerPoint presentation file.
			Dim result As String = "SetBrightnessAndTransparencyForGradientStop.pptx"

			'Save the document
			presentation.SaveToFile(result, FileFormat.Pptx2010)

			'  Release resources
			presentation.Dispose()

			' Launch the saved file
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace