Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace ToSpecificSizeImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Conversion.pptx")

			'Save the first slide to Image and set the image size to 600*400
			Dim img As Image = ppt.Slides(0).SaveAsImage(600, 400)

			'Save image to file
			Dim result As String = "ToSpecificSizeImage.png"
			img.Save(result, System.Drawing.Imaging.ImageFormat.Png)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            using (var images = ppt.Slides[0].SaveAsImage(600, 400))
'            {
'                FileStream fileStream = new FileStream("ToSpecificSizeImage.png", FileMode.Create, FileAccess.Write);
'                images.CopyTo(fileStream);
'                fileStream.Flush();
'                images.Dispose();
'            }
'            

			PresentationDocViewer(result)
		End Sub

	Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
	End Sub
	End Class
End Namespace