Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace SetMasterBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Set the slide background of master
			presentation.Masters(0).SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom
			presentation.Masters(0).SlideBackground.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			presentation.Masters(0).SlideBackground.Fill.SolidColor.Color = Color.LightSalmon

			'Save the document
			Dim result As String = "SetSlideMasterBackground_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub
	End Class
End Namespace