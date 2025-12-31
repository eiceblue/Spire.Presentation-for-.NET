Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace ToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToImage.pptx")

			'Save PPT document to images
			For i As Integer = 0 To presentation.Slides.Count - 1
				Dim fileName As String = String.Format("ToImage-img-{0}.png", i)
				Dim image As Image = presentation.Slides(i).SaveAsImage()
				image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)
				System.Diagnostics.Process.Start(fileName)
			Next i

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            for (int i = 0; i < presentation.Slides.Count; i++)
'            {
'                using (var images = presentation.Slides[i].SaveAsImage())
'                {
'                    String fileName = String.Format("ToImage_img_{0}.png", i);
'                    FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
'                    images.CopyTo(fileStream);
'                    fileStream.Flush();
'                    images.Dispose();
'                }
'            }
'            

			 presentation.Dispose()

		End Sub
	End Class
End Namespace