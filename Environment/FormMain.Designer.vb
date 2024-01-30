<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.txtInfo = New System.Windows.Forms.TextBox()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.txtArea = New System.Windows.Forms.TextBox()
        Me.lblArea = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtInfo
        '
        Me.txtInfo.Location = New System.Drawing.Point(12, 41)
        Me.txtInfo.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtInfo.Multiline = True
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.ReadOnly = True
        Me.txtInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtInfo.Size = New System.Drawing.Size(831, 434)
        Me.txtInfo.TabIndex = 0
        Me.txtInfo.TabStop = False
        '
        'txtPath
        '
        Me.txtPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPath.Location = New System.Drawing.Point(65, 496)
        Me.txtPath.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.ReadOnly = True
        Me.txtPath.Size = New System.Drawing.Size(659, 20)
        Me.txtPath.TabIndex = 1
        Me.txtPath.TabStop = False
        '
        'lblPath
        '
        Me.lblPath.AutoSize = True
        Me.lblPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPath.Location = New System.Drawing.Point(8, 499)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(51, 13)
        Me.lblPath.TabIndex = 2
        Me.lblPath.Text = "File Path:"
        '
        'btnCopy
        '
        Me.btnCopy.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCopy.Location = New System.Drawing.Point(754, 493)
        Me.btnCopy.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(89, 25)
        Me.btnCopy.TabIndex = 3
        Me.btnCopy.TabStop = False
        Me.btnCopy.Text = "Select All"
        Me.btnCopy.UseVisualStyleBackColor = True
        '
        'txtArea
        '
        Me.txtArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArea.Location = New System.Drawing.Point(85, 13)
        Me.txtArea.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtArea.Name = "txtArea"
        Me.txtArea.ReadOnly = True
        Me.txtArea.Size = New System.Drawing.Size(108, 20)
        Me.txtArea.TabIndex = 5
        Me.txtArea.TabStop = False
        '
        'lblArea
        '
        Me.lblArea.AutoSize = True
        Me.lblArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArea.Location = New System.Drawing.Point(12, 16)
        Me.lblArea.Name = "lblArea"
        Me.lblArea.Size = New System.Drawing.Size(67, 13)
        Me.lblArea.TabIndex = 6
        Me.lblArea.Text = "Namespace:"
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(890, 538)
        Me.Controls.Add(Me.lblArea)
        Me.Controls.Add(Me.txtArea)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.lblPath)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.txtInfo)
        Me.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "FormMain"
        Me.Text = "Environment"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtInfo As TextBox
    Friend WithEvents txtPath As TextBox
    Friend WithEvents lblPath As Label
    Friend WithEvents btnCopy As Button
    Friend WithEvents txtArea As TextBox
    Friend WithEvents lblArea As Label
End Class
