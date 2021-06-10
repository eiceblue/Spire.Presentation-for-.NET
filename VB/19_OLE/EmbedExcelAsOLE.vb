Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace EmbedExcelAsOLE
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Presentaion document
			Dim ppt As New Presentation()

			'Load a image file
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\EmbedExcelAsOLE.png")
			Dim oleImage As IImageData = ppt.Images.Append(image)
			Dim rec As New Rectangle(80, 60, image.Width, image.Height)

			'Insert an OLE object to presentation based on the Excel data
			Dim oleObject As Spire.Presentation.IOleObject = ppt.Slides(0).Shapes.AppendOleObject("excel", File.ReadAllBytes("..\..\..\..\..\..\Data\EmbedExcelAsOLE.xlsx"), rec)
			oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage
			oleObject.ProgId = "Excel.Sheet.12"

			'Save the document
			ppt.SaveToFile("EmbedExcelAsOLE_result.pptx", Spire.Presentation.FileFormat.Pptx2010)
			Process.Start("EmbedExcelAsOLE_result.pptx")
		End Sub
	End Class
End Namespace
