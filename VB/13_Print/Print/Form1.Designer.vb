Namespace Print
	Partial Public Class Form1
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
			Me.pbLogo = New PictureBox()
			Me.lblDescription = New Label()
			Me.btnRun = New Button()
			CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' pbLogo
			' 
			Me.pbLogo.Image = (CType(resources.GetObject("pbLogo.Image"), Image))
			Me.pbLogo.ImeMode = Windows.Forms.ImeMode.NoControl
			Me.pbLogo.Location = New Point(15, 15)
			Me.pbLogo.Name = "pbLogo"
			Me.pbLogo.Size = New Size(56, 48)
			Me.pbLogo.SizeMode = Windows.Forms.PictureBoxSizeMode.AutoSize
			Me.pbLogo.TabIndex = 69
			Me.pbLogo.TabStop = False
			' 
			' lblDescription
			' 
			Me.lblDescription.Font = New Font("Verdana", 8.25F)
			Me.lblDescription.ImeMode = Windows.Forms.ImeMode.NoControl
			Me.lblDescription.Location = New Point(74, 15)
			Me.lblDescription.Name = "lblDescription"
			Me.lblDescription.Size = New Size(358, 79)
			Me.lblDescription.TabIndex = 70
			Me.lblDescription.Text = resources.GetString("lblDescription.Text")
			' 
			' btnRun
			' 
			Me.btnRun.ImeMode = Windows.Forms.ImeMode.NoControl
			Me.btnRun.Location = New Point(360, 97)
			Me.btnRun.Name = "btnRun"
			Me.btnRun.Size = New Size(75, 23)
			Me.btnRun.TabIndex = 71
			Me.btnRun.Text = "Run"
			Me.btnRun.UseVisualStyleBackColor = True
'			Me.btnRun.Click += New System.EventHandler(Me.btnRun_Click)
			' 
			' Form1
			' 
			Me.AutoScaleDimensions = New SizeF(6F, 12F)
			Me.AutoScaleMode = Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New Size(452, 136)
			Me.Controls.Add(Me.btnRun)
			Me.Controls.Add(Me.lblDescription)
			Me.Controls.Add(Me.pbLogo)
			Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
			Me.MaximizeBox = False
			Me.MinimizeBox = False
			Me.Name = "Form1"
			Me.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
			Me.Text = "Print"
			CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub

		#End Region

		Private pbLogo As PictureBox
		Private lblDescription As Label
		Private WithEvents btnRun As Button
	End Class
End Namespace