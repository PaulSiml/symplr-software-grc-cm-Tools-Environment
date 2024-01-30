Imports System.IO
Imports System.Management

Public Class FormMain

    Dim line As String = ""
    Dim instance As String = ""
    Dim version As String = ""

    Dim dir As String = ""
    Dim rtn As String = ""

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If File.Exists(gbLog) Then
            File.Delete(gbLog)
        End If

        Dim value As String = ReadRegistry(Microsoft.Win32.Registry.LocalMachine, "Software\WOW6432Node\InterSystems\IRIS\Configurations\IRISHealth\Directory", "")

        If Not value = "" And Not value.Contains("does not exist") Then

            instance = "Iris"
            dir = value
            version = ReadRegistry(Microsoft.Win32.Registry.LocalMachine, "Software\WOW6432Node\InterSystems\IRIS\Configurations\IRISHealth", "Version")

        End If


        If instance = "" Then

            value = ReadRegistry(Microsoft.Win32.Registry.LocalMachine, "Software\WOW6432Node\InterSystems\Cache\Configurations\CACHE\Directory", "")

            If Not value = "" And Not value.Contains("does not exist") Then

                instance = "Cache"
                dir = value
                version = ReadRegistry(Microsoft.Win32.Registry.LocalMachine, "Software\WOW6432Node\InterSystems\Cache\Configurations\CACHE", "Version")

            End If

        End If

        Dim frmLogin As New FormLogin

        frmLogin.MPath = dir
        frmLogin.Response = False

        frmLogin.ShowDialog()

        If frmLogin.Response = False Then
            Me.Close()
            End
        End If

        Dim myFont As New Font(FontFamily.GenericMonospace, 8)
        txtPath.Font = myFont

        txtPath.Text = gbLog
        txtArea.Text = IIf(gbArea = "", "None", gbArea)

    End Sub

    Private Sub LoadData()

        Dim qt As String = Chr(34)

        DisplaySection("Loading ...")

        LogIt("----- OPERATING SYSTEM -----")

        DisplaySection("Operating System")
        Dim pgm As String = "cmd.exe"
        Dim parms As String = "/c systeminfo.exe | findstr /B /C:""OS Name"" /B /C:""OS Version"" /B /C:""Processor(s)"" /B /C:""Total Physical Memory"" >>" & gbLog
        Shell(pgm, parms, True)


        LogIt("----- VOLUME -----")

        DisplaySection("Volume")
        pgm = "PowerShell.exe"
        parms = "Get-Volume | Out-File -Append -Encoding ASCII -FilePath " & gbLog
        Shell(pgm, parms, True)


        LogIt("----- INSTANCE -----")

        DisplaySection("Instance")
        line = "Type: " & instance
        LogIt(line)

        line = "Version: " & version
        LogIt(line)

        line = "Install Path: " & dir
        LogIt(line)
        LogIt("License: (Below)")
        LogIt(" ")

        ' license key info
        callM(Mumps("d ^CKEY w !"), "", "%SYS", rtn)


        LogIt("----- MIDAS -----")

        DisplaySection("Midas")
        callM(Mumps("s lbl=""Area: "",gbArea=" & qt & gbArea & qt & " w lbl_$j("""", 25-$l(lbl))_gbArea,!"), "", gbArea, rtn)
        callM(Mumps("s lbl=""Side Code: "" w lbl_$j("""", 25-$l(lbl))_$G(^MIDIC(2000)),!"), "", gbArea, rtn)
        callM(Mumps("s lbl=""Version: "" w lbl_$j("""", 25-$l(lbl))_$G(^MIDIC(2000,""MIDAS+"")),!"), "", gbArea, rtn)
        callM(Mumps("s lbl=""Deployment Patch Level: "" w lbl_$j("""", 25-$l(lbl))_$G(^MIDIC(2000,""MIDAS+"",""DEPLOYMENT PATCH LEVEL"")),!"), "", gbArea, rtn)
        callM(Mumps("s lbl=""Live to Test Copy: "" w lbl_$j("""", 25-$l(lbl))_$s($G(^MIDIC(2000,""MIDAS+"",""LIVE TO TEST COPY""))'="""":^(""LIVE TO TEST COPY""),1:""None""),!"), "", gbArea, rtn)


        LogIt("----- DATABASES -----")

        Dim mps As String = "[RETURN]k properties s status=##Class(Security.Services).Get(""%Service_Telnet"",.properties) s RETURN=properties(""Enabled"")"
        Dim status As Boolean = callM(mps, "", "%SYS", rtn)
        Dim enbl As String = rtn

        If status = True Then
            If enbl = 0 Then
                ' turn on TelNet (if not on)
                mps = "k properties s status=##Class(Security.Services).Get(""%Service_Telnet"",.properties) s properties(""Enabled"")=1,status=##Class(Security.Services).Modify(""%Service_Telnet"",.properties)"
                callM(mps, "", "%SYS", rtn)
            End If
        End If

        DisplaySection("Databases")
        ' this section requires utility input - thus a terminal script that requires TelNet

        Dim fileName As String = Path.Combine(Application.StartupPath, "input.scr")
        pgm = Path.Combine(dir, IIf(dir.Contains("IRIS"), "bin\iristerm.exe", "bin\cterm.exe"))
        parms = " /console=cn_iptcp:127.0.0.1[23] " & fileName & " " & gbUserName & " " & gbPassword & " " & gbArea
        BuildInputScript()
        Shell(pgm, parms, True)
        File.Delete(fileName)

        If enbl = 0 Then
            ' turn off TelNet (if off initially)
            mps = "k properties s status=##Class(Security.Services).Get(""%Service_Telnet"",.properties) s properties(""Enabled"")=0,status=##Class(Security.Services).Modify(""%Service_Telnet"",.properties)"
            callM(mps, "", "%SYS", rtn)
        End If


        LogIt("----- SYSTEM EMAIL -----")

        DisplaySection("System Email")
        callM(Mumps("d INIT^MPUMAIL,DISP^MPUMAIL w !"), "", gbArea, rtn)


        LogIt("----- INTERFACES -----")

        DisplaySection("Interfaces")
        callM(Mumps("d ^MIDPAR,SRVLIST^%ZMSERVER w !"), "", gbArea, rtn)

        LogIt("----- Inbound -----")
        callM(Mumps("zw ^MIDIC(""INBOUND"")"), "", gbArea, rtn)

        LogIt("----- Outbound -----")
        callM(Mumps("zw ^MIDIC(""OUTBOUND"")"), "", gbArea, rtn)

        Sleep(2000)

        RemoveBlankLines(gbLog)

        ' load info into UI
        txtInfo.Text = ReadFile(gbLog)


    End Sub

    Private Function ReadFile(ByVal infoFile As String) As String

        Dim line As String = ""
        Dim str As String = ""

        ReadFile = ""

        If File.Exists(infoFile) = False Then Exit Function

        Dim fsStream As New System.IO.FileStream(infoFile, System.IO.FileMode.Open, System.IO.FileAccess.Read, FileShare.Read)

        Try
            ' create a stream reader attach the file stream to it
            Dim fsReader As New System.IO.StreamReader(fsStream)
            ' read lines until EOF is reached.

            Do
                line = fsReader.ReadLine()
                If line Is Nothing Then Exit Do
                ' get namespace section
                str = IIf(str = "", line, str & vbCrLf & line)

            Loop
            ' close the writer when done
        Catch err As Exception
            str = ""
        Finally
            ' close the file when done
            fsStream.Close()
        End Try


        ReadFile = str

    End Function

    Private Function Mumps(ByVal exe As String) As String

        Dim qt As String = Chr(34)
        Dim writeLog As String = " s gbLog=" & qt & gbLog & qt & " o gbLog:(:/WRITE:/APP):3 u gbLog x pgm c gbLog"
        ' add quotes for executable mumps variable
        exe = exe.Replace("""", """""")
        ' create mumps code to be run by Distirbution.Scrips CallM
        Mumps = "s pgm=" & qt & exe & qt & writeLog

    End Function

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click

        Clipboard.Clear()

        txtInfo.Focus()
        txtInfo.SelectAll()

        Clipboard.SetText(txtInfo.SelectedText)

    End Sub

    Private Sub FormMain_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Sleep(500)
        LoadData()

    End Sub

    Private Sub DisplaySection(ByVal text As String)

        ' display section loading in text box
        txtInfo.Text = IIf(txtInfo.Text = "", text, txtInfo.Text & vbCrLf & text)

    End Sub

End Class
