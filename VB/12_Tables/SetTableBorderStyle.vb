Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace SetTableBorderStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_1.pptx")

			'Find the table by looping through all the slides, and then set borders for it. 
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is ITable Then
						For Each row As TableRow In (TryCast(shape, ITable)).TableRows
							For Each cell As Cell In row
								cell.BorderTop.FillType = FillFormatType.Solid
								cell.BorderBottom.FillType = FillFormatType.Solid
								cell.BorderLeft.FillType = FillFormatType.Solid
								cell.BorderRight.FillType = FillFormatType.Solid
							Next cell
						Next row
					End If
				Next shape
			Next slide

			Dim result As String = "Result-SetTableBorderStyle.pptx"

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