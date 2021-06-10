Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace ExtractTextFromSmartArt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ExtractTextFromSmartArt.pptx")

			'Traverse through all the slides of the PPT file and find the SmartArt shapes.
			Dim st As New StringBuilder()
		   st.AppendLine("Below is extracted text from SmartArt:")
			For i As Integer = 0 To presentation.Slides.Count - 1
				For j As Integer = 0 To presentation.Slides(i).Shapes.Count - 1
					If TypeOf presentation.Slides(i).Shapes(j) Is ISmartArt Then
						Dim smartArt As ISmartArt = TryCast(presentation.Slides(i).Shapes(j), ISmartArt)

						'Extract text from SmartArt and append to the StringBuilder object.
						For k As Integer = 0 To smartArt.Nodes.Count - 1
							st.AppendLine(smartArt.Nodes(k).TextFrame.Text)
						Next k
					End If
				Next j
			Next i

			Dim result As String = "Result-ExtractTextFromSmartArt.txt"

			'Save to file.
			File.WriteAllText(result, st.ToString())

			'Launch the file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace