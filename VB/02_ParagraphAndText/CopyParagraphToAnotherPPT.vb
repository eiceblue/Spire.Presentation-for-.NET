Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace CopyParagraphToAnotherPPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the source file
			Dim ppt1 As New Presentation()
			ppt1.LoadFromFile("..\..\..\..\..\..\Data\TextTemplate.pptx")

			'Get the text from the first shape on the first slide
			Dim sourceshp As IShape = ppt1.Slides(0).Shapes(0)
			Dim text As String = (CType(sourceshp, IAutoShape)).TextFrame.Text

			'Load the target file
			Dim ppt2 As New Presentation()
			ppt2.LoadFromFile("..\..\..\..\..\..\Data\CopyParagraph.pptx")

			'Get the first shape on the first slide from the target file
			Dim destshp As IShape = ppt2.Slides(0).Shapes(0)

			'Add the text to the target file
			CType(destshp, IAutoShape).TextFrame.Text += vbLf & vbLf & text

			'Save the document
			Dim result As String = "CopyParagraphToAnotherPPT.pptx"
			ppt2.SaveToFile(result, FileFormat.Pptx2013)
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