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

        'encrypy the document with password "test"
        presentation.Encrypt("test")

        'save the document
        presentation.SaveToFile("encrypt.pptx", FileFormat.Pptx2007)
        System.Diagnostics.Process.Start("encrypt.pptx")

    End Sub
End Class