Imports Spire.Presentation

Namespace ConvertPPSToPPTX
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Conversion.pps")

			'Save the PPS document to PPTX file format
			Dim result As String = "ConvertPPSToPPTX.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			'Launch and view the resulted PPTX file
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