Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace InsertImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\InsertImage.pptx")

			'Insert image to PPT
			Dim ImageFile2 As String = "..\..\..\..\..\..\Data\InsertImage.png"
			Dim rect1 As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 280, 140, 120, 120)
			Dim image As IEmbedImage = presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile2, rect1)
			image.Line.FillType = FillFormatType.None

			'Save the document
			presentation.SaveToFile("InsertImage.pptx", FileFormat.Pptx2010)
			Process.Start("InsertImage.pptx")
		End Sub
	End Class
End Namespace