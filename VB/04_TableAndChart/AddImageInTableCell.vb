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
			'create a PPT document
			Dim presentation As New Presentation()

			'create a table and set the table style
			Dim widths() As Double = { 100, 100 }
			Dim heights() As Double = { 100, 100 }
			Dim table As ITable = presentation.Slides(0).Shapes.AppendTable(100, 80, widths, heights)
			table.StylePreset = TableStylePreset.LightStyle1Accent2

			'load the image and insert it into table
			Dim imgData As IImageData = presentation.Images.Append(Image.FromFile("..\..\..\..\..\..\Data\flower.jpg"))
			table(0, 0).FillFormat.FillType = FillFormatType.Picture
			table(0, 0).FillFormat.PictureFill.Picture.EmbedImage = imgData
			table(0, 0).FillFormat.PictureFill.FillType = PictureFillType.Stretch

			'save the document
			presentation.SaveToFile("InsertImageInTable.pptx", FileFormat.Pptx2010)
			Process.Start("InsertImageInTable.pptx")
		End Sub
	End Class
End Namespace
