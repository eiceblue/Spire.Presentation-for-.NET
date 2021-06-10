Imports Spire.Presentation

Namespace ReplaceText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim tagValues As New Dictionary(Of String, String)()
			tagValues.Add("Spire.Presentation for .NET", "Spire.PPT")

			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\TextTemplate.pptx")

			ReplaceTags(ppt.Slides(0), tagValues)

			'Save the document
			Dim result As String = "ReplaceText.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub ReplaceTags(ByVal pSlide As ISlide, ByVal TagValues As Dictionary(Of String, String))
			For Each curShape As IShape In pSlide.Shapes
				If TypeOf curShape Is IAutoShape Then
					For Each tp As TextParagraph In (TryCast(curShape, IAutoShape)).TextFrame.Paragraphs
						For Each curKey In TagValues.Keys
							If tp.Text.Contains(curKey) Then
								tp.Text = tp.Text.Replace(curKey, TagValues(curKey))
							End If
						Next curKey
					Next tp
				End If
			Next curShape
		End Sub
	End Class
End Namespace