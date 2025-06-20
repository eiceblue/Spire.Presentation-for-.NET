Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetColumnsCountOfTextFrame
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\ColumnsCount.pptx")

			'Get the first shape in first slide and set column count of text for it.
			Dim shape1 As IAutoShape = CType(ppt.Slides(0).Shapes(0), IAutoShape)
			shape1.TextFrame.ColumnCount = 3

			'Get the second shape in second slide and set column count of text for it.
			Dim shape2 As IAutoShape = CType(ppt.Slides(1).Shapes(0), IAutoShape)
			shape2.TextFrame.ColumnCount = 2

			'Save the document
			ppt.SaveToFile("result.pptx", FileFormat.Pptx2010)
			Process.Start("result.pptx")
		End Sub
	End Class
End Namespace


