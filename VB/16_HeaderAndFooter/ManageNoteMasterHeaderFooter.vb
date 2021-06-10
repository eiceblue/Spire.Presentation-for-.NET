Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ManageNoteMasterHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			Dim loadPath As String = "..\..\..\..\..\..\Data\PPTHasHeader.pptx"
			Dim savePath As String = "ManageNoteMasterHeaderFooter.pptx"

			'Load presentation
			presentation.LoadFromFile(loadPath)

			'Set the note Masters header and footer
			Dim noteMasterSlide As INoteMasterSlide = presentation.NotesMaster
			If Not CType(noteMasterSlide, Object).Equals(Nothing) Then
				For Each shape As Shape In noteMasterSlide.Shapes
					If Not shape.Placeholder.Equals(Nothing) Then
						If shape.Placeholder.Type.Equals(PlaceholderType.Header) Then
							TryCast(shape, IAutoShape).TextFrame.Text = "change the header by Spire"
						End If
						If shape.Placeholder.Type.Equals(PlaceholderType.Footer) Then
							TryCast(shape, IAutoShape).TextFrame.Text = "change the footer by Spire"
						End If
					End If
				Next shape
			End If

			presentation.SaveToFile(savePath, FileFormat.Pptx2013)
			Process.Start(savePath)
		End Sub
	End Class
End Namespace