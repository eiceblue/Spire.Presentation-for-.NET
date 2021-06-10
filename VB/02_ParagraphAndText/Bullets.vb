Imports Spire.Presentation.Drawing
Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation


Namespace Bullets
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Bullets.pptx")

			Dim shape As IAutoShape = CType(presentation.Slides(0).Shapes(1), IAutoShape)

			For Each para As TextParagraph In shape.TextFrame.Paragraphs
				'Add the bullets
				para.BulletType = TextBulletType.Numbered
				para.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod

			Next para

			'Save the document and launch
			presentation.SaveToFile("bullets.pptx", FileFormat.Pptx2010)
			Process.Start("bullets.pptx")
		End Sub
	End Class
End Namespace