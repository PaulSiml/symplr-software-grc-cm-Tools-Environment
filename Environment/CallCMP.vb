Module CallCMP
    Public Function callM(ByVal routine As String, ByVal parms As String, ByVal area As String, ByRef sReturn As String) As Boolean

        ' connect to caché/iris

        Dim factory As CacheActiveX.Factory = CreateObject("CacheActiveX.Factory")

        Dim mObj As Object
        Dim connectstring As String

        callM = False

        Try
            ' Create instance of factory object
            If Not factory.IsConnected() Then
                ' Establish connection to server
                connectstring = "cn_iptcp:127.0.0.1[" & gbPort & "]:" & area & ":" & gbUserName & ":" & gbPassword

                ' no connection?
                If factory.Connect(connectstring) = False Then
                    sReturn = "Caché ActiveX Factory Connection Failure"
                    Exit Function
                End If
                '
            End If

            ' static class to be used in mumps environment

            mObj = factory.Static("Distribution.Script")

            Select Case routine

                ' site parameters
                Case "GetSP" : sReturn = mObj.GetSP(parms, area)
                Case "SetSP" : sReturn = mObj.SetSP(Piece(parms, ",", 1), area, Piece(parms, ",", 2, 99))
                ' globals
                Case "GetGlb" : sReturn = mObj.GetGlb(area, parms)
                Case "SetGlb" : sReturn = mObj.SetGlb(area, parms)

                Case Else

                    ' default to run routine / execute command
                    sReturn = mObj.callM(routine, parms)

            End Select
            '
            callM = True

        Catch err As Exception

            MessageBox.Show("Area: " & area & vbCrLf & vbCrLf & err.Message, "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            ' disconnect
            If factory.IsConnected() Then factory.Disconnect()
        End Try

    End Function

End Module
