Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports System.Collections
Imports System.ComponentModel
Imports System.Text

Namespace RemoveUnusedLayoutMaster
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load document from disk
			Dim ppt As New Presentation()
			ppt.LoadFromFile("../../../../../../Data/PPTSample_1.pptx")

			'Create an array list
			Dim list As New List(Of IActiveSlide)()
			For i As Integer = 0 To ppt.Slides.Count - 1
				'Get the layout used by slide
				Dim layout As IActiveSlide = CType(ppt.Slides(i).Layout, IActiveSlide)
				list.Add(layout)
			Next i

			'Loop through masters and layouts
			For i As Integer = 0 To ppt.Masters.Count - 1
				Dim masterlayouts As IMasterLayouts = ppt.Masters(i).Layouts
				For j As Integer = masterlayouts.Count - 1 To 0 Step -1
					If Not list.Contains(CType(masterlayouts(j), IActiveSlide)) Then
						'Remove unused layout
						masterlayouts.RemoveMasterLayout(j)
					End If
				Next j
			Next i

			'Save the document
			Dim outputFile As String = "RemoveUnusedLayoutMaster_out.pptx"
			ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
			ppt.Dispose()
			Process.Start(outputFile)
		End Sub
	End Class
End Namespace
