Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Data.SqlTypes
Imports System.IO
Imports System.Text


Namespace ToSvgUsingGraphicTag
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ExtractImage.pptx")
			'When saving a PPT document to SVG, save the graphics in the PPT document as image tags
			presentation.SaveToSvgOption.ConvertPictureUsingGraphicTag = True
			For i As Integer = 0 To presentation.Slides.Count - 1
				Dim fileName As String = String.Format("ToSVG-{0}.svg", i)
				Dim fs As New FileStream(fileName, FileMode.Create)
				'Convert the  slide to SVG
				Dim silde() As Byte = presentation.Slides(i).SaveToSVG()
				fs.Write(silde, 0, silde.Length)
				Process.Start(fileName)
			Next i

		End Sub
	End Class
End Namespace
