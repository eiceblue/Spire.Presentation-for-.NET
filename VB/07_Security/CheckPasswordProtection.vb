Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace CheckPasswordProtection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Check whether a PPT document is password protected
			Dim isProtected As Boolean=presentation.IsPasswordProtected("..\..\..\..\..\..\Data\Template_Ppt_4.pptx")

			'Show the result by message box
			MessageBox.Show("The file is " & (If(isProtected, "password ", "not password ")) & "protected!")
		End Sub
	End Class
End Namespace