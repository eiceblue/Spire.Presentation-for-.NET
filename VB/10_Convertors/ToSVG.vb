Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text
Imports System.Windows.Forms
Imports System.IO

Public Class Form1

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        'create PPT document
        Dim presentation As New Presentation()

        'load PPT file from disk
        presentation.LoadFromFile("..\..\..\..\..\..\Data\source.pptx")

        Dim svgBytes As Queue(Of Byte()) = presentation.SaveToSVG()
        For i As Integer = 0 To svgBytes.Count - 1
            Dim fs As FileStream = New FileStream(String.Format("{0}.svg", i), FileMode.Create)
            Dim bytes As Byte() = svgBytes.Dequeue()
            fs.Write(bytes, 0, bytes.Length)
        Next

    End Sub
End Class