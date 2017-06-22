Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Public Class Form1

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        'create PPT document
        Dim presentation As New Presentation()

        'load PPT file from disk
        presentation.LoadFromFile("..\..\..\..\..\..\Data\source.pptx")

        'save PPT document to images
        For i As Integer = 0 To presentation.Slides.Count - 1
            Dim fileName As [String] = [String].Format("result-img-{0}.png", i)
            Dim image As Image = presentation.Slides(i).SaveAsImage()
            image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)
            System.Diagnostics.Process.Start(fileName)
        Next

    End Sub
End Class