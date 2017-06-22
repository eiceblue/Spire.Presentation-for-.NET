Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation.Drawing

Public Class Form1

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        'create PPT document
        Dim presentation As New Presentation()

        'load PPT file from disk
        presentation.LoadFromFile("..\..\..\..\..\..\Data\edit.pptx")

        'edit the first shape
        Dim shape As IAutoShape = DirectCast(presentation.Slides(0).Shapes(0), IAutoShape)
        Dim para As New TextParagraph()
        para.Text = "Edit Sample"
        para.TextRanges(0).LatinFont = New TextFont("Myriad Pro")
        para.TextRanges(0).FontHeight = 24
        shape.TextFrame.Paragraphs.Append(para)

        'save the document
        presentation.SaveToFile("edited.pptx", FileFormat.Pptx2007)
        System.Diagnostics.Process.Start("edited.pptx")

    End Sub
End Class