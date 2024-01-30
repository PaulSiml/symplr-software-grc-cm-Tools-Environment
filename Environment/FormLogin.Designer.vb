<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormLogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.chkShow = New System.Windows.Forms.CheckBox()
        Me.txtArea = New System.Windows.Forms.TextBox()
        Me.lblArea = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.PasswordLabel = New System.Windows.Forms.Label()
        Me.UsernameLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'chkShow
        '
        Me.chkShow.AutoSize = True
        Me.chkShow.Location = New System.Drawing.Point(187, 77)
        Me.chkShow.Name = "chkShow"
        Me.chkShow.Size = New System.Drawing.Size(53, 17)
        Me.chkShow.TabIndex = 19
        Me.chkShow.Text = "Show"
        Me.chkShow.UseVisualStyleBackColor = True
        '
        'txtArea
        '
        Me.txtArea.Location = New System.Drawing.Point(13, 119)
        Me.txtArea.Name = "txtArea"
        Me.txtArea.Size = New System.Drawing.Size(119, 20)
        Me.txtArea.TabIndex = 16
        Me.txtArea.Tag = ""
        '
        'lblArea
        '
        Me.lblArea.Location = New System.Drawing.Point(11, 99)
        Me.lblArea.Name = "lblArea"
        Me.lblArea.Size = New System.Drawing.Size(220, 23)
        Me.lblArea.TabIndex = 11
        Me.lblArea.Text = "&Update Area"
        Me.lblArea.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(40, 166)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(92, 23)
        Me.btnCancel.TabIndex = 18
        Me.btnCancel.Text = "&Cancel"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(140, 166)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(92, 23)
        Me.btnOK.TabIndex = 17
        Me.btnOK.Text = "&OK"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(12, 75)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(169, 20)
        Me.txtPassword.TabIndex = 15
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(14, 32)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(220, 20)
        Me.txtUserName.TabIndex = 14
        '
        'PasswordLabel
        '
        Me.PasswordLabel.Location = New System.Drawing.Point(10, 55)
        Me.PasswordLabel.Name = "PasswordLabel"
        Me.PasswordLabel.Size = New System.Drawing.Size(220, 23)
        Me.PasswordLabel.TabIndex = 12
        Me.PasswordLabel.Text = "&Password"
        Me.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UsernameLabel
        '
        Me.UsernameLabel.Location = New System.Drawing.Point(12, 12)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(220, 23)
        Me.UsernameLabel.TabIndex = 13
        Me.UsernameLabel.Text = "&User name"
        Me.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FormLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(257, 211)
        Me.Controls.Add(Me.chkShow)
        Me.Controls.Add(Me.txtArea)
        Me.Controls.Add(Me.lblArea)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.PasswordLabel)
        Me.Controls.Add(Me.UsernameLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormLogin"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Environment - Instance Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkShow As CheckBox
    Friend WithEvents txtArea As TextBox
    Friend WithEvents lblArea As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnOK As Button
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents txtUserName As TextBox
    Friend WithEvents PasswordLabel As Label
    Friend WithEvents UsernameLabel As Label
End Class
