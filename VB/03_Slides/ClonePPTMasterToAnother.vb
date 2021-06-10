Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace ClonePPTMasterToAnother
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load PPT1 from disk
			Dim presentation1 As New Presentation()
			presentation1.LoadFromFile("..\..\..\..\..\..\Data\CloneMaster1.pptx")

			'Load PPT2 from disk
			Dim presentation2 As New Presentation()
			presentation2.LoadFromFile("..\..\..\..\..\..\Data\CloneMaster2.pptx")

			'Add masters from PPT1 to PPT2
			For Each masterSlide As IMasterSlide In presentation1.Masters
				presentation2.Masters.AppendSlide(masterSlide)
			Next masterSlide

			'Save the document
			Dim result As String = "ClonePPTMasterToAnother.pptx"
			presentation2.SaveToFile(result, FileFormat.Pptx2013)

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