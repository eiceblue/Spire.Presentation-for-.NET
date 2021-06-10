Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace MultipleLevelBullets
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Bullets2.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Access the first placeholder in the slide and typecasting it as AutoShape
			Dim tf1 As ITextFrameProperties = (CType(slide.Shapes(1), IAutoShape)).TextFrame

			'Access the first Paragraph and set bullet style
			Dim para As TextParagraph= tf1.Paragraphs(0)
			para.BulletType = TextBulletType.Symbol
			para.BulletChar = Convert.ToChar(8226)
			para.Depth = 0

			 'Access the second Paragraph and set bullet style
			 para = tf1.Paragraphs(1)
			 para.BulletType = TextBulletType.Symbol
			 para.BulletChar = "-"c
			 para.Depth = 1

			 'Access the third Paragraph and set bullet style
			 para = tf1.Paragraphs(2)
			 para.BulletType = TextBulletType.Symbol
			 para.BulletChar = Convert.ToChar(8226)
			 para.Depth = 2

			 'Access the fourth Paragraph and set bullet style
			 para = tf1.Paragraphs(3)
			 para.BulletType = TextBulletType.Symbol
			 para.BulletChar = "-"c
			 para.Depth = 3

			Dim result As String = "MultipleLevelBullets_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)
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