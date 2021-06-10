Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AutoFitTextOrShape
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

			'Set the AutofitType property to Shape
			Dim textShape2 As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(150, 100, 150, 80))
			textShape2.TextFrame.Text = "Resize shape to fit text."
			textShape2.TextFrame.AutofitType = TextAutofitType.Shape

			'Set the AutofitType property to Normal
			Dim textShape1 As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(400, 100, 150, 80))
			textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape."
			textShape1.TextFrame.AutofitType = TextAutofitType.Normal

			'Save the document
			Dim result As String = "AutoFitTextOrShape.pptx"
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