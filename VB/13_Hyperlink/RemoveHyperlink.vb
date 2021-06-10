Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace RemoveHyperlink
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

			'Get the shape and its text with hyperlink.
			Dim shape As IAutoShape = TryCast(presentation.Slides(0).Shapes(0), IAutoShape)

			'Set the ClickAction property into null to remove the hyperlink.
			shape.TextFrame.TextRange.ClickAction = Nothing

			Dim result As String = "Result-RemoveHyperlink.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
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