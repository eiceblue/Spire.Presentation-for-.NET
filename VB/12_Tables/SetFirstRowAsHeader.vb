Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetFirstRowAsHeader
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim loadPath As String = "..\..\..\..\..\..\Data\NormalTable.pptx"
			Dim savePath As String = "SetFirstRowAsHeader.pptx"
			Dim table As ITable = Nothing

			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile(loadPath)

			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ITable Then
					table = TryCast(shape, ITable)
				End If

			Next shape
			table.FirstRow = True

			'Save the file
			presentation.SaveToFile(savePath, FileFormat.Pptx2010)
			Process.Start(savePath)
		End Sub
	End Class
End Namespace