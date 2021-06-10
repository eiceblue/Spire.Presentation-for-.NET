Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace ExtractOLEObject
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ExtractOLEObject.pptx")

			'Loop through the slides and shapes
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is IOleObject Then
						'Find OLE object
						Dim oleObject As IOleObject = TryCast(shape, IOleObject)

						'Get its data and write to file
						Dim bytes() As Byte = oleObject.Data
						Select Case oleObject.ProgId
							Case "Excel.Sheet.8"
								File.WriteAllBytes("result.xls", bytes)
							Case "Excel.Sheet.12"
								File.WriteAllBytes("result.xlsx", bytes)
							Case "Word.Document.8"
								File.WriteAllBytes("result.doc", bytes)
							Case "Word.Document.12"
								File.WriteAllBytes("result.docx", bytes)
							Case "PowerPoint.Show.8"
								File.WriteAllBytes("result.ppt", bytes)
							Case "PowerPoint.Show.12"
								File.WriteAllBytes("result.pptx", bytes)
						End Select
					End If
				Next shape
			Next slide
			MessageBox.Show("Completed!")
		End Sub
	End Class
End Namespace