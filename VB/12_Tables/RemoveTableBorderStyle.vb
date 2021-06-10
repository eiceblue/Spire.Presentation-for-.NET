Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace RemoveTableBorderStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is ITable Then
						For Each row As TableRow In (TryCast(shape, ITable)).TableRows
							For Each cell As Cell In row
								cell.BorderTop.FillType = FillFormatType.None
								cell.BorderBottom.FillType = FillFormatType.None
								cell.BorderLeft.FillType = FillFormatType.None
								cell.BorderRight.FillType = FillFormatType.None
							Next cell
						Next row
					End If
				Next shape
			Next slide

			Dim result As String = "Result-RemoveTableBorderStyle.pptx"

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