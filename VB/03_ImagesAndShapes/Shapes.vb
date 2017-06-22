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

        'set background Image
        Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
        Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
        presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
        presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

        'append new shape - Triangle
        Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Triangle, New RectangleF(50, 100, 100, 100))

        'set the color and fill style of shape
        shape.Fill.FillType = FillFormatType.Solid
        shape.Fill.SolidColor.Color = Color.LightGreen
        shape.ShapeStyle.LineColor.Color = Color.White

        'append new shape - Ellipse
        shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Ellipse, New RectangleF(270, 100, 150, 100))
        shape.ShapeStyle.LineColor.Color = Color.White

        'append new shape - FivePointedStar
        shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.FivePointedStar, New RectangleF(50, 270, 150, 150))

        'set the color of shape
        shape.Fill.FillType = FillFormatType.Gradient
        shape.Fill.SolidColor.Color = Color.Black
        shape.ShapeStyle.LineColor.Color = Color.White

        'append new shape - Rectangle
        shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(300, 300, 100, 120))

        'set the color of shape
        shape.Fill.FillType = FillFormatType.Solid
        shape.Fill.SolidColor.Color = Color.Tomato
        shape.ShapeStyle.LineColor.Color = Color.Tomato

        'append new shape - BentUpArrow
        shape = presentation.Slides(0).Shapes.AppendShape(ShapeType.BentUpArrow, New RectangleF(500, 300, 150, 100))

        'set the color of shape
        shape.Fill.FillType = FillFormatType.Gradient
        shape.Fill.Gradient.GradientStops.Append(1.0F, KnownColors.Olive)
        shape.Fill.Gradient.GradientStops.Append(0, KnownColors.PowderBlue)
        shape.ShapeStyle.LineColor.Color = Color.White

        'save the document
        presentation.SaveToFile("shape.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("shape.pptx")

    End Sub
End Class