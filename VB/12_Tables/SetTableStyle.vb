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
			'Creat a ppt document and load file
			Dim ppt As New Presentation()
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SetTableStyle.pptx")

			'Get tbe table
			Dim table As ITable = Nothing
			For Each shape As IShape In ppt.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = CType(shape, ITable)

					'Set the table style from TableStylePreset and apply it to selected table
					table.StylePreset = TableStylePreset.MediumStyle1Accent2
				End If
			Next shape
			'Save the file
			ppt.SaveToFile("SetTableStyle_result.pptx", FileFormat.Pptx2010)
			Process.Start("SetTableStyle_result.pptx")
		End Sub
	End Class
End Namespace
