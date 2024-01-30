Imports System.IO
Imports System.Threading
Imports Microsoft.Win32

Module Utilities

    Public gbLog As String = Path.Combine(Application.StartupPath, "System_Info.txt")
    Public gbUserName As String = ""
    Public gbPassword As String = ""
    Public gbArea As String = ""
    Public gbPort As String = "1972"
    Public gbWebPort As String = "57772"

    Public Function ReadRegistry(ByVal parentKey As RegistryKey, ByVal subKey As String, ByVal valueName As String) As String

        Dim key As RegistryKey
        Dim parent As String, msg As String

        ReadRegistry = ""

        Try
            'Open the registry key
            key = parentKey.OpenSubKey(subKey, True)
            If key Is Nothing Then 'if the key doesn't exist
                parent = parentKey.Name.ToString
                msg = "Reg Key: " & parent & "\" & subKey & " does not exist "
                ReadRegistry = msg
                Exit Try
            End If
            'Get the value
            ReadRegistry = CStr(key.GetValue(valueName))
        Catch err As Exception

        End Try

    End Function
    Public Function Piece(ByVal target As String, ByVal delim As String, ByVal position As Integer, Optional ByVal last As Integer = -9999) As String

        '  Input parameters : target = the string to get a piece of
        '                     delim  = the delimiter sing
        '                     piece  = the piece to get
        '----------------------------------------------------------------------------------------------
        Piece = ""

        Dim values As String() = target.Split(delim)

        ' default means no last piece request
        If last = -9999 Then last = position

        ' prevent loop execution
        If last < position Or (position < 1) Then position = 999

        For I = (position - 1) To values.GetUpperBound(0)

            If Piece = "" Then
                Piece = values(I)
            Else
                Piece = Piece & delim & values(I)
            End If

            If I >= (last - 1) Then Exit For

        Next

    End Function

    Public Function Shell(ByVal fileName As String, ByVal args As String, ByVal hidden As Boolean) As Boolean
        ' run a command & wait until it's done
        '
        ' fileName is the complete path including the file to be run
        Dim process As System.Diagnostics.Process
        Shell = 0
        Try
            process = New System.Diagnostics.Process
            process.StartInfo.FileName = fileName
            process.StartInfo.Arguments = args

            If hidden = True Then
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            Else
                process.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            End If

            process.StartInfo.WorkingDirectory = Path.GetDirectoryName(fileName)
            If fileName.Contains(".exe") Then process.StartInfo.Verb = "runas"
            If fileName.Contains(".bat") Then process.StartInfo.Verb = "runas"

            process.Start()

            'Wait until the process passes back an exit code 
            process.WaitForExit()

            'Free resources associated with this process
            process.Close()
            Shell = 1
        Catch
            'MessageBox.Show("Could not start process " & fileName, "Error")
        End Try

    End Function

    Public Sub Sleep(ByVal mSec As Integer)

        Dim cnt As Integer

        Do
            Thread.Sleep(100)
            Application.DoEvents()
            cnt = cnt + 100
            If cnt >= mSec Then Exit Do

        Loop

    End Sub
    Public Function ReadCpf(ByVal cpf As String) As String

        Dim line As String = ""
        Dim ignore As String = ""
        Dim name As String = ""
        Dim nspstr As String = ""

        ReadCpf = ""

        If File.Exists(cpf) = False Then Exit Function

        Dim nspSection As Boolean = False
        Dim fsStream As New System.IO.FileStream(cpf, System.IO.FileMode.Open, System.IO.FileAccess.Read, FileShare.Read)


        Try
            ' create a stream reader attach the file stream to it
            Dim fsReader As New System.IO.StreamReader(fsStream)
            ' read lines until EOF is reached.
            ' search for Namespace section
            '
            ' example data from cpf file
            '   [Namespaces]
            '   %SYS=CACHESYS
            '   DOCBOOK = DOCBOOK
            '   MAA = PAPP
            '   MAT = TRAINAPP
            '   MAV = TESTAPP
            '   SAMPLES = SAMPLES
            '   USER = USER
            '
            ' [Namespaces]
            '   %SYS=IRISSYS
            '   HSCUSTOM = HSCUSTOM
            '   HSLIB = HSLIB
            '   HSSYS = HSSYS
            '   MAA = PAPP
            '   MAT = TRAINAPP
            '   MAV = TESTAPP
            '   USER = USER
            '   %ALL = %DEFAULTDB

            ignore = "%SYS,%ALL,HSCUSTOM,HSLIB,HSSYS,DOCBOOK,SAMPLES,USER"
            Do
                line = fsReader.ReadLine()
                If line Is Nothing Then Exit Do
                ' get namespace section
                If line.StartsWith("[Namespaces]") Then
                    nspSection = True
                    ' go to next line
                    line = fsReader.ReadLine()
                End If

                If nspSection = True Then
                    If line = "" Then
                        ' stop reading section
                        nspSection = False
                    Else
                        ' we are in the namespaces section
                        If line.Contains("=") Then
                            name = Piece(line, "=", 1).Trim
                            ' we only want Midas additional datasets
                            If InStr(ignore, name) = 0 Then
                                nspstr = IIf(nspstr = "", name, nspstr & "," & name)

                            End If
                        End If
                    End If
                End If

            Loop
            ' close the writer when done
        Catch err As Exception

        Finally
            ' close the file when done
            fsStream.Close()
        End Try

        ReadCpf = nspstr

    End Function

    Public Sub LogIt(ByVal line As String, Optional ByVal fileName As String = "")

        If fileName = "" Then fileName = gbLog

        If fileName = "" Then Exit Sub

        Dim writer As New StreamWriter(fileName, True)
        writer.WriteLine(line)
        writer.Close()

    End Sub

    Public Sub BuildInputScript()


        Dim fileName As String = Path.Combine(Application.StartupPath, "input.scr")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        LogIt("echo: On", fileName)
        LogIt("on error: $Missing", fileName)
        LogIt("timer: 3", fileName)
        LogIt("wait for:Do you agree to abide by these policies (Y/<N>) :", fileName)
        LogIt("send: Y<CR>", fileName)
        LogIt("wait for:Username", fileName)
        LogIt("$Missing:", fileName)
        LogIt("timer: 0", fileName)
        LogIt("send: <p1><CR>", fileName)
        LogIt("wait for:Password", fileName)
        LogIt("send: <p2><CR>", fileName)
        LogIt("send: znspace ""%SYS""<CR>", fileName)
        LogIt("send: do ALL^%FREECNT<CR>", fileName)
        LogIt("send: D:\temp\System_Info.txt<CR>", fileName)
        LogIt("send: EWA<CR>", fileName)
        LogIt("send: halt<CR>", fileName)
        LogIt("$skiptag:", fileName)
        LogIt("pause: 50", fileName)
        LogIt("terminate", fileName)

    End Sub

    Public Sub RemoveBlankLines(ByVal fileName As String)

        Dim lines() As String = File.ReadAllLines(fileName)
        Dim newLines As New ArrayList
        Dim formatLines As New ArrayList

        For Each line As String In lines
            If line <> "" Then
                newLines.Add(line)
            End If
        Next


        For Each line As String In newLines

            formatLines.Add(line)
            If line.StartsWith("----- ") Then
                'add line after section header
                formatLines.Add("")
                Dim index As Integer = formatLines.IndexOf(line)
                ' add line prior to section header
                formatLines.Insert(index, "")

            End If
        Next

        Dim displayLines() As String = formatLines.ToArray(GetType(String))

        File.WriteAllLines(gbLog, displayLines)


    End Sub

End Module
