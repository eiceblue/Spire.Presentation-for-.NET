Imports Spire.Presentation
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace ExtractImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ExtractImage.pptx")

			For i As Integer = 0 To ppt.Images.Count - 1
				Dim ImageName As String = String.Format("Images-{0}.png", i)
				'Extract image
				Dim image As Image = ppt.Images(i).Image
				image.Save(ImageName)

				'////////////////Use the following code for netstandard dlls/////////////////////////
'                
'                SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(ppt.Images[i].Image);
'                FileStream fileStream = new FileStream(ImageName, FileMode.Create, FileAccess.Write);
'                image.Encode().SaveTo(fileStream);
'                fileStream.Flush();
'                image.Dispose();
'                          
			Next i
			ppt.Dispose()
		End Sub
	End Class
End Namespace
