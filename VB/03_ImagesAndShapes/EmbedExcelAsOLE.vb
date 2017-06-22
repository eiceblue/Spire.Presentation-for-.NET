Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.IO
Imports System.Text
'Imports System.Threading.Tasks

Namespace EmbedExcelAsOLE
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            Dim image As Image = System.Drawing.Image.FromFile("..\..\..\..\..\..\Data\EmbedExcelAsOLE.png")
			Dim ppt As New Presentation()
			Dim oleImage As IImageData = ppt.Images.Append(image)
			Dim rec As New Rectangle(60, 60, image.Width, image.Height)

			'insert an OLE object to presentation based on the Excel data

			Dim oleObject As Spire.Presentation.IOleObject = ppt.Slides(0).Shapes.AppendOleObject("excel", File.ReadAllBytes("..\..\..\..\..\..\Data\DatatableSample.xlsx"), rec)
			oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage
			oleObject.ProgId = "Excel.Sheet.12"


			'save the document
			ppt.SaveToFile("InsertOle.pptx", Spire.Presentation.FileFormat.Pptx2010)
			Process.Start("InsertOle.pptx")
		End Sub
	End Class
End Namespace
