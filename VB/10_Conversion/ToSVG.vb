Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.Text
Imports System.IO
Imports Spire.Presentation

Namespace ToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			'Create PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ToSVG.pptx")

			'Retain note when converting a PPT document to SVG files
			presentation.IsNoteRetained = True

			Dim svgBytes As Queue(Of Byte())=presentation.SaveToSVG()
			Dim count As Integer = svgBytes.Count
            For i As Integer = 0 To count - 1
                Dim bt() As Byte = svgBytes.Dequeue()
                Dim fileName As String = String.Format("ToSVG-{0}.svg", i)
                Dim fs As New FileStream(fileName, FileMode.Create)
                fs.Write(bt, 0, bt.Length)
                Process.Start(fileName)
            Next i
        End Sub

    End Class
End Namespace