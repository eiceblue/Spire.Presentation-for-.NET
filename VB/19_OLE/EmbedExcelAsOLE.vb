Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Namespace EmbedExcelAsOLE
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a Presentaion document
			Dim ppt As New Presentation()

			'Load a image file
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\EmbedExcelAsOLE.png")
			Dim oleImage As IImageData = ppt.Images.Append(image)

			'////////////////Use the following code for netstandard dlls///////////////////////// 
'            
'            FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\EmbedExcelAsOLE.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fileStream.Length];
'            fileStream.Read(bytes, 0, bytes.Length);
'            fileStream.Close();
'            Stream stream = new MemoryStream(bytes);          
'            IImageData oleImage = ppt.Images.Append(stream);
'            stream.Close();
'            fileStream.Close();
'            SkiaSharp.SKBitmap image = SkiaSharp.SKBitmap.Decode(@"..\..\..\..\..\..\Data\EmbedExcelAsOLE.png");
'            

			Dim rec As New Rectangle(80, 60, image.Width, image.Height)

			'Insert an OLE object to presentation based on the Excel data
			Dim oleObject As Spire.Presentation.IOleObject = ppt.Slides(0).Shapes.AppendOleObject("excel", File.ReadAllBytes("..\..\..\..\..\..\Data\EmbedExcelAsOLE.xlsx"), rec)
			oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage
			oleObject.ProgId = "Excel.Sheet.12"

			'Save the document
			ppt.SaveToFile("EmbedExcelAsOLE_result.pptx", Spire.Presentation.FileFormat.Pptx2010)
			System.Diagnostics.Process.Start("EmbedExcelAsOLE_result.pptx")
		End Sub
	End Class
End Namespace
