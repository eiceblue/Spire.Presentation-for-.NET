Imports Spire.Presentation
Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Namespace ExportFormulas
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Load the PPT file that contains formulas
            Dim presentation As Presentation = New Presentation()
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExportFormulas.pptx")

            ' Initialize a StringBuilder to store all exported formulas
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim count As Integer = 0

            ' Iterate through all slides and shapes to export every formula
            Dim i As Integer = 0
            While i < presentation.Slides.Count
                Dim slide As ISlide = presentation.Slides[i]

                Dim j As Integer = 0
                While j < slide.Shapes.Count
                    ' Attempt to cast the shape to an AutoShape that can contain text
                    Dim autoShape As IAutoShape = slide.Shapes[j] as IAutoShape

                    ' Skip shapes that are not AutoShapes or have no text frame
                    If autoShape == null || autoShape.TextFrame == null Then
                        Continue
                    End If

                    ' Iterate through all paragraphs in the text frame
                    Dim k As Integer = 0
                    While k < autoShape.TextFrame.Paragraphs.Count
                        Dim paragraph As TextParagraph = autoShape.TextFrame.Paragraphs[k]

                        Try
                            ' Export the formula from the paragraph in both MathML and LaTeX formats
                            Dim mathML As String = paragraph.ExportMathML()
                            Dim latex As String = paragraph.ExportLaTex()

                            ' Skip if the paragraph does not contain any formula
                            If string.IsNullOrEmpty(mathML) && string.IsNullOrEmpty(latex) Then
                                Continue
                            End If

                            ' Increment the formula counter and append the exported data
                            count += 1
                            stringBuilder.AppendLine("Formula " + count + ":")
                            stringBuilder.AppendLine("  MathML:")
                            stringBuilder.AppendLine("  " + mathML)
                            stringBuilder.AppendLine("  LaTeX:")
                            stringBuilder.AppendLine("  " + latex)
                            stringBuilder.AppendLine()
        Catch
                            ' Skip the paragraphs that contain no formula or have export errors
                            ' Continue processing remaining paragraphs
                        End Try
                    Next
                Next
            Next

            ' Prepend a summary line indicating the total number of formulas exported
            stringBuilder.Insert(0, "Exported " + count + " formula(s) in total." + Environment.NewLine + Environment.NewLine)

            ' Save all the exported formulas into a text file
            File.WriteAllText("ExportFormulas_out.txt", stringBuilder.ToString())

            ' Launch the text file to display the exported formulas
            System.Diagnostics.Process.Start("ExportFormulas_out.txt")
        End Sub

    End Class
End Namespace
