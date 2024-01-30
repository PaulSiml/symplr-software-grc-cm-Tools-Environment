Imports System.Windows.Forms

Imports System.IO

Public Class FormLogin

    Property Response As Boolean
    Property MPath As String

    Property Login As Boolean = False

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        ' login to Cache/Iris - test communication with simple line of Mumps code to %SYS

        gbUserName = txtUserName.Text
        gbPassword = txtPassword.Text
        gbArea = txtArea.Text

        Login = False

        Me.Hide()

        If gbArea <> "" Then

            If Not (",MAA,MAT,MAV,MAB,MAC,MAD,MAE,MAF,MAG,MBV,MCV,MDV,MEV,MFV,MGV,").Contains("," & gbArea & ",") Then

                If MessageBox.Show("Namespace name is not a standard Midas name." & vbCrLf & vbCrLf & "Continue?", "Validation Failure", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2) = DialogResult.No Then Exit Sub

            End If

        End If

        If MPath <> "" Then

            ' determine if we have communication - execute line of mumps code in %SYS & Update Area

            Dim rtn As String = ""
            ' check for Distribution.Scripts in both area and %SYS
            Dim status As Boolean = callM("S A=1", "", "%SYS", rtn)
            If status Then status = callM("S A=1", "", gbArea, rtn)

            Login = status

            If status = False Then


                If MessageBox.Show("Login Failure - No Communication With Cache/Iris!!" & vbCrLf & vbCrLf & "1) Incorrect Username/Password" & vbCrLf & vbCrLf & "2) Invalid Area" & vbCrLf & vbCrLf & "3) Instance may not be running" & vbCrLf & vbCrLf & "4) Distribution.Script is missing from Update Area" & vbCrLf & vbCrLf & "5) Distribution.Script is missing from %SYS" & vbCrLf & vbCrLf & "Continue?", "Environment - Connection Failure", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2) = DialogResult.No Then Exit Sub

            End If

        End If


        Response = True


        Me.Close()

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Response = False
        Me.Close()

    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txtUserName.Text = gbUserName
        txtPassword.Text = gbPassword
        txtArea.Text = gbArea

    End Sub


    Private Sub chkShow_CheckedChanged(sender As Object, e As EventArgs) Handles chkShow.CheckedChanged
        txtPassword.PasswordChar = If(chkShow.Checked, "", "*")
    End Sub

End Class

