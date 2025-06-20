Imports System.IO
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace GetDefaultTextFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim inputFile As String = "..\..\..\..\..\..\Data\GetDefaultTextFormat.pptx"
			Dim outputFile As String = "GetDefaultTextFormat.txt"

			' Create Presentation object and load the file
			Dim presentation As New Presentation()
			presentation.LoadFromFile(inputFile)

			' Get the first shape of the first slide
			Dim shape As IAutoShape = TryCast(presentation.Slides(0).Shapes(0), IAutoShape)

			' Get the display format of the text in shape
			Dim format As DefaultTextRangeProperties = shape.TextFrame.Paragraphs(0).TextRanges(0).DisPlayFormat

			' Determine whether the format is bold or italic
			File.AppendAllText(outputFile, "Is the first shape text bolded :" & format.IsBold & vbCrLf)
			File.AppendAllText(outputFile, "Is the first shape text italicized :" & format.IsItalic & vbCrLf)

			' Dispose
			presentation.Dispose()

			Process.Start(outputFile)

		End Sub

	End Class
End Namespace