Imports Spire.Presentation
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace SaveChartAsImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a PPT document 
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SaveChartAsImage.pptx")

			'Save chart as image in .png format
			Dim image As Image = presentation.Slides(0).Shapes.SaveAsImage(0)
			image.Save("Chart_result.png", System.Drawing.Imaging.ImageFormat.Png)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            System.IO.Stream stream = ppt.Slides[0].SaveAsImage();
'            byte[] buff = new byte[stream.Length];
'            stream.Read(buff, 0, buff.Length);
'            FileStream fs = new FileStream("Chart_result.png", FileMode.Create);
'            fs.Write(buff);
'            fs.Close();
'			

			System.Diagnostics.Process.Start("Chart_result.png")
		End Sub
	End Class
End Namespace
