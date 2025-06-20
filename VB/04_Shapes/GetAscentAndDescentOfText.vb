Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace GetAscentAndDescentOfText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Presentation object
			Dim ppt As New Presentation()

			' Load a PowerPoint file from a specified location
			ppt.LoadFromFile("..\..\..\..\..\..\..\Data\GetAscentAndDescentOfText.pptx")

			' Create a StringBuilder to store text information
			Dim builder As New StringBuilder()

			' Access the first slide in the presentation
			Dim slide As ISlide = ppt.Slides(0)

			' Access the first AutoShape in the slide
			Dim autoshape As IAutoShape = TryCast(slide.Shapes(0), IAutoShape)

			' Retrieve the layout lines from the TextFrame of the AutoShape
			Dim lines As IList(Of LineText) = autoshape.TextFrame.GetLayoutLines()

			' Iterate through each layout line
			For i As Integer = 0 To lines.Count - 1
				' Get the ascent and descent properties of the current line
				Dim ascent As Single = lines(i).Ascent
				Dim descent As Single = lines(i).Descent

				' Append information about the line, ascent, and descent to the StringBuilder
				builder.AppendLine("lines" & i & vbTab & "ascent: " & ascent & vbTab & "descent: " & descent)
			Next i

			' Specify the name for the result file
			Dim result As String = "GetAscentAndDescentOfText.txt"

			' Save to the text file
			File.WriteAllText(result, builder.ToString())

			' Dispose of the Presentation object to release resources
			ppt.Dispose()

			' Launch the result file
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace