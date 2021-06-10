Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ModifyHyperlink
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

			'Find the hyperlinks you want to edit.
			Dim shape As IAutoShape = CType(presentation.Slides(0).Shapes(0), IAutoShape)

			'Edit the link text and the target URL.
			shape.TextFrame.TextRange.ClickAction.Address = "http://www.e-iceblue.com"
			shape.TextFrame.TextRange.Text = "E-iceblue"

			Dim result As String = "Result-ModifyHyperlink.pptx"

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