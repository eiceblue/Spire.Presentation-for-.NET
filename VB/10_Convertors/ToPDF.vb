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
        presentation.LoadFromFile("..\..\..\..\..\..\Data\Presentation1.pptx")

        'save the PPT do PDF file format
        presentation.SaveToFile("ToPdf.pdf", FileFormat.PDF)
        System.Diagnostics.Process.Start("ToPdf.pdf")

    End Sub
End Class