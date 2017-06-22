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

        'load the PPT file from disk
        presentation.LoadFromFile("..\..\..\..\..\..\Data\source97.ppt")

        'save the PPT document to PPTX file format
        presentation.SaveToFile("ToPPTX.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("ToPPTX.pptx")

    End Sub
End Class