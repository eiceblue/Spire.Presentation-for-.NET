Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace DisableAdvanceAfterTime
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a Presentation object
			Dim ppt As New Presentation()

			' Load the PPT file from the specified path
			ppt.LoadFromFile("..\..\..\..\..\..\..\Data\DisableAdvanceAfterTime.pptx")

			' Get the first slide and disable the selected advance after time setting
			ppt.Slides(0).SlideShowTransition.SelectedAdvanceAfterTime = False

			' Specify the name for the output PowerPoint presentation file.
			Dim result As String = "output.pptx"

			' Save the modified PPT to the specified path
			ppt.SaveToFile(result, FileFormat.Pptx2013)

			' Dispose of the Presentation object to free up resources
			ppt.Dispose()

			' Launch the saved file
			PresentationDocViewer(result)
		End Sub

		Private Shared Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace