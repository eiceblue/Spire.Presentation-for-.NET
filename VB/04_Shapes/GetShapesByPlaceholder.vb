Imports System.Collections
Imports System.IO
Imports Spire.Presentation
Imports Spire.Presentation.Drawing

Namespace GetShapesByPlaceholder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()
			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\Data\GetShapesByPlaceholder.pptx")
			'Get Placeholder
			Dim placeholder As Placeholder = ppt.Slides(1).Shapes(0).Placeholder
			'Get Shapes by Placeholder
			Dim shapes() As IShape = ppt.Slides(1).GetPlaceholderShapes(placeholder)
			Dim text As String = ""
			'Iterate over all the shapes
			For i As Integer = 0 To shapes.Length - 1
				'If shape is IAutoShape
				If TypeOf shapes(i) Is IAutoShape Then
					Dim autoShape As IAutoShape = TryCast(shapes(i), IAutoShape)
					If autoShape.TextFrame IsNot Nothing Then
						text &= autoShape.TextFrame.Text & vbCrLf

					End If
				End If
			Next i
			Dim result As String = "GetShapesByPlaceholder_output.txt"
			File.WriteAllText(result, text)

			'Launch the PowerPoint file
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