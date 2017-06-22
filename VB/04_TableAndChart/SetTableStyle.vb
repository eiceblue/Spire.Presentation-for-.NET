Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace SetTableStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'creat a ppt document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\test.pptx")

			'get tbe table
			Dim table As ITable = TryCast(ppt.Slides(0).Shapes(0), ITable)

			'set the table style from TableStylePreset and apply it to selected table
			table.StylePreset = TableStylePreset.DarkStyle1Accent6

			'save the file
			ppt.SaveToFile("tableStyle.pptx", FileFormat.Pptx2010)
			Process.Start("tableStyle.pptx")
		End Sub
	End Class
End Namespace
