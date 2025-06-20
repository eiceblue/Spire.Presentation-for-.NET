Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace GetSlideLayoutName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\GetSlideLayoutName.pptx")

			Dim builder As New StringBuilder()

			'Loop through the slides of PPT document
			For i As Integer = 0 To presentation.Slides.Count - 1
				'Get the name of slide layout
				Dim name As String = presentation.Slides(i).Layout.Name
				builder.AppendLine(String.Format("The name of slide {0} layout is: {1}", i,name))
			Next i

			'Save the document
			Dim result As String="GetSlideLayoutName_out.txt"
			File.WriteAllText(result, builder.ToString())

			'Launch the Pdf file
			DocumentViewer(result)
		End Sub
		Private Sub DocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try

		End Sub
	End Class
End Namespace
