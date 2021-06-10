Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ToPdfWithSpecificPageSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToPDF.pptx")

			'Set A4 page size
			presentation.SlideSize.Type = SlideSizeType.A4

			'Set landscape orientation
			presentation.SlideSize.Orientation = SlideOrienation.Landscape

			Dim result As String = "ToPdfWithSpecifiedPageSize_result.pdf"
			'Save the PPT to PDF file format
			presentation.SaveToFile(result, FileFormat.PDF)

			Viewer(result)
		End Sub

		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace