Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace AddImageInTableCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\AddImageInTableCell.pptx")

			'Get the first shape
			Dim table As ITable = TryCast(ppt.Slides(0).Shapes(0), ITable)

			'Load the image and insert it into table cell
			Dim pptImg As IImageData = ppt.Images.Append(Image.FromFile("..\..\..\..\..\..\Data\PresentationIcon.png"))

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'			FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\PresentationIcon.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fileStream.Length];
'            fileStream.Read(bytes, 0, bytes.Length);
'            fileStream.Close();
'            Stream stream = new MemoryStream(bytes);
'            IImageData pptImg = ppt.Images.Append(stream);
'            stream.Close();
'            fileStream.Close();
'            

			table(1, 1).FillFormat.FillType = FillFormatType.Picture
			table(1, 1).FillFormat.PictureFill.Picture.EmbedImage = pptImg
			table(1, 1).FillFormat.PictureFill.FillType = PictureFillType.Stretch

			'Save the document
			ppt.SaveToFile("AddImageInTableCell_result.pptx", FileFormat.Pptx2010)
			System.Diagnostics.Process.Start("AddImageInTableCell_result.pptx")
		End Sub
	End Class
End Namespace
