Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetRightToLeftColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\TwoColumns.pptx")

			'Get the second shape
			Dim shape As IAutoShape = TryCast(ppt.Slides(0).Shapes(1), IAutoShape)
			'Set columns style to right-to-left
			shape.TextFrame.RightToLeftColumns = True

			'Save the document
			Dim result As String = "SetRightToLeftColumns.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
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