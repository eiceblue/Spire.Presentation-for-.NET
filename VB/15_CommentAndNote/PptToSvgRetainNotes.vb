Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace PptToSvgRetainNotes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_5.pptx")

			'Retain the notes while converting PowerPoint file to svg file.
			presentation.IsNoteRetained = True

			'Convert presentation slides to svg file.
			Dim bytes As Queue(Of Byte()) = presentation.SaveToSVG()

			Dim length As Integer = bytes.Count
			For i As Integer = 0 To length - 1
				Dim result As String = String.Format("output_{0}.svg", i)
				Dim filestream As New FileStream(result, FileMode.Create)
				Dim outputBytes() As Byte = bytes.Dequeue()
				filestream.Write(outputBytes, 0, outputBytes.Length)

				'Launch the PowerPoint file.
				PptDocumentViewer(result)
			Next i

			presentation.Dispose()
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace