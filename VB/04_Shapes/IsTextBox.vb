Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace IsTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\IsTextboxSample.pptx")

			Dim builder As New StringBuilder()

			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is IAutoShape Then
						'Judge if the shape is textbox
						Dim isTextbox As Boolean = shape.IsTextBox
						builder.AppendLine(If(isTextbox, "shape is text box", "shape is not text box"))
					End If
				Next shape
			Next slide

			'Write the content of builder to txt file
			File.WriteAllText("IsTextbox.txt", builder.ToString())
			Process.Start("IsTextbox.txt")
		End Sub
	End Class
End Namespace