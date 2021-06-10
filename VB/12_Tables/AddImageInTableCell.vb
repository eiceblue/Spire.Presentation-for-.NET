Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text

Namespace AddImageInTableCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\AddImageInTableCell.pptx")

			'Get the first shape
			Dim table As ITable = TryCast(ppt.Slides(0).Shapes(0), ITable)

			'Load the image and insert it into table cell
			Dim pptImg As IImageData = ppt.Images.Append(Image.FromFile("..\..\..\..\..\..\Data\PresentationIcon.png"))

			table(1, 1).FillFormat.FillType = FillFormatType.Picture
			table(1, 1).FillFormat.PictureFill.Picture.EmbedImage = pptImg
			table(1, 1).FillFormat.PictureFill.FillType = PictureFillType.Stretch

			'Save the document
			ppt.SaveToFile("AddImageInTableCell_result.pptx", FileFormat.Pptx2010)
			Process.Start("AddImageInTableCell_result.pptx")
		End Sub
	End Class
End Namespace
