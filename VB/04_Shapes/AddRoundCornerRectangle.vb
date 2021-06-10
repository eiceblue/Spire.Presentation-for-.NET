Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace AddRoundCornerRectangle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Set background image
			Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
			Dim rect As New RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
			ppt.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
			ppt.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

			'Append a round corner rectangle and set its radius
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendRoundRectangle(300, 90, 100, 200, 80)
			'Set the color and fill style of shape
			shape.Fill.FillType = FillFormatType.Solid
			shape.Fill.SolidColor.Color = Color.LightBlue
			shape.ShapeStyle.LineColor.Color = Color.SkyBlue
			'Rotate the shape to 90 degree
			shape.Rotation = 90

			'Save the document
			Dim result As String = "AddRoundCornerRectagle.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
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