Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Text.Encoding
Imports System.Text.RegularExpressions

Public Class classFTP

    Private _ServerType As String = ""
    Private _EventLog As New StringBuilder

    Private MyFTPSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    Private MyFTPEntry As IPEndPoint
    Private Loggedinto As Boolean = False
    Private PasvSocket As Socket

    Private bNewResponse As Boolean
    Private LastResponse As New classLastResponse(0, "")

    Public Directorys As New ArrayList
    Public Files As New ArrayList

    Public Event OnErrorMsg(ByVal ErrorMsg As String)
    Public Event OnEventAdded(ByVal MsgNo As Integer, ByVal Msg As String)
    Public Event UploadProgress(ByVal Filename As String, ByVal BytesUploaded As Integer, ByVal Filesize As Integer)
    Public Event DownloadProgress(ByVal Filename As String, ByVal BytesReveived As Integer, ByVal Filesize As Integer)
    Public Event OnDisconnected()
    Public Event OnConnection()

#Region "Regex Expressions for phasing directorys"

    Private regex_UNIX_1 As New Regex("(?<dir>[\-dr])(?<permission>([\-r][\-w][\-xs]){3})\s+\d+\s+\w+\s+\w+\s+(?<size>\d+)\s+(?<timestamp>\w+\s+\d+\s+\d{4})\s+(?<name>.+)")
    Private regex_UNIX_2 As New Regex("(?<dir>[\-dr])(?<permission>([\-r][\-w][\-xs]){3})\s+\d+\s+\d+\s+(?<size>\d+)\s+(?<timestamp>\w+\s+\d+\s+\d{4})\s+(?<name>.+)")
    Private regex_UNIX_3 As New Regex("(?<dir>[\-dr])(?<permission>([\-r][\-w][\-xs]){3})\s+\d+\s+\d+\s+(?<size>\d+)\s+(?<timestamp>\w+\s+\d+\s+\d{2}:\d{2})\s+(?<name>.+)")
    Private regex_WIND_1 As New Regex("(?<timestamp>\d{2}\-\d{2}\-\d{2}\s+\d{2}:\d{2}[Aa|Pp][mM])\s+(?<dir>\<\w+\>){0,1}(?<size>\d+){0,1}\s+(?<name>.+)")

#End Region

#Region "Properties"

    Public ReadOnly Property ServerType() As String
        Get
            Return _ServerType
        End Get
    End Property

    Public ReadOnly Property Eventlog() As String
        Get
            Return _EventLog.ToString
        End Get
    End Property

    Public ReadOnly Property Connected() As Boolean
        Get
            If (MyFTPSocket.Connected = True) And (Loggedinto = True) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

#End Region

#Region "Classes"

    Private Class classStatus
        Public Buffer(1024) As Byte
        Public ServerResponseMessage As String
    End Class

    Private Class classGetFile
        Public Buffer(1024) As Byte
        Public ServerResponseMessage As String
        Public RemoteFilename As String
        Public LocalFilename As String
        Public ReceivedBytes As Integer
        Public FileSize As Integer
        Public Socket As Socket
        Public LocalFile As IO.FileStream
    End Class

    Private Class classLastResponse
        Private _MsgNo As Integer
        Private _Msg As String
        Public ReadOnly Property MsgNo() As Integer
            Get
                Return _MsgNo
            End Get
        End Property
        Public ReadOnly Property Msg() As String
            Get
                Return _Msg
            End Get
        End Property
        Public Sub New(ByVal MsgNo As Integer, ByVal Msg As String)
            _MsgNo = MsgNo
            _Msg = Msg
        End Sub
    End Class

    Public Class classFileAttb
        Private _filename As String
        Private _size As Integer
        Private _Timestamp As Date
        Private _isdirectory As Boolean
        Public Sub New(ByVal Filename As String, ByVal Size As Integer, ByVal Timestamp As Date, ByVal IsDirectory As Boolean)
            _filename = Filename
            _size = Size
            '_Timestamp = Timestamp
            _isdirectory = IsDirectory
        End Sub
        Public ReadOnly Property Filename() As String
            Get
                Return _filename
            End Get
        End Property
        Public ReadOnly Property Size() As Integer
            Get
                Return _size
            End Get
        End Property
        'Public ReadOnly Property Timestamp() As Date
        '    Get
        '        Return _Timestamp
        '    End Get
        'End Property
        Public ReadOnly Property IsDirectory() As Boolean
            Get
                Return _isdirectory
            End Get
        End Property
    End Class

#End Region

#Region "Start of connection"

    Public Function Connect(ByVal Hostname As String, ByVal Username As String, ByVal Password As String, Optional ByVal Directory As String = "", Optional ByVal Port As Integer = 21) As Boolean
        Try
            If ConnectToHost(Hostname, Port) Then
                ' Start event handler
                GetResponse()
                If WaitForResponseNo(220) = False Then
                    Disconnect()
                    RaiseEvent OnErrorMsg("Login failed - No welcome message")
                    Return False
                End If
                Send("USER " & Username)
                If WaitForResponseNo(331) = False Then
                    Disconnect()
                    RaiseEvent OnErrorMsg("Login failed")
                    Return False
                End If
                'Console.WriteLine("Login")
                Send("PASS " & Password)
                If WaitForResponseNo(230) = False Then
                    Disconnect()
                    RaiseEvent OnErrorMsg("Login failed - Login failed")
                    Return False
                End If
                'Console.WriteLine("PASS")
                Send("SYST")
                If WaitForResponseNo(215) = False Then
                    Disconnect()
                    RaiseEvent OnErrorMsg("Login failed - No system identiforcation")
                    Return False
                End If
                'Console.WriteLine("SYST")
                _ServerType = LastResponse.Msg
                If WaitForLogin() = False Then Exit Function
                'Console.WriteLine("Logged in")
                ' Reset the response watcher
                ResetResponse()
                If Len(Directory) = 0 Then ChangeDirectory(CurrentDirectory) Else ChangeDirectory(Directory)
                RaiseEvent OnConnection()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function ConnectToHost(ByVal Hostname As String, ByVal Port As Integer) As Boolean
        Try
            If IsNumeric(Replace(Hostname, ".", "")) Then
                MyFTPEntry = New IPEndPoint(Hostname, Port)
            Else
                Dim MyHost As IPHostEntry = Dns.GetHostByName(Hostname)
                MyFTPEntry = New IPEndPoint(MyHost.AddressList(0), Port)
            End If
            MyFTPSocket.Connect(MyFTPEntry)
            Return True
        Catch
            RaiseEvent OnErrorMsg("Error connecting to host: " & Hostname)
            Return False
        End Try
    End Function

#End Region

#Region "End of Connection"

    Public Sub Disconnect()
        StartsocketShutdown()
    End Sub

    Private Function StartsocketShutdown()
        MyFTPSocket.Shutdown(SocketShutdown.Both)
        MyFTPSocket.Close()
        ' Create new socket to use ftp control again
        MyFTPSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        ' Raise the event for disconnected
        Loggedinto = False
        RaiseEvent OnDisconnected()
    End Function

#End Region

#Region "Send Message / And public commands"

    Public Function Send(ByVal Command As String)
        Try
            If MyFTPSocket.Connected Then
                Dim SendBuffer As Byte() = ASCII.GetBytes(Command & Chr(13) & Chr(10))
                MyFTPSocket.Send(SendBuffer, SendBuffer.Length, 0)
                _EventLog.Append(Command & vbCrLf)
                'RaiseEvent OnEventAdded(0, Command & vbCrLf)
            Else
                Throw New Exception("Server not connected")
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Public Function Delete(ByVal Filename As String) As Boolean
        Try
            ResetResponse()
            Send("DELE " & Filename)
            If Not (WaitForResponse.MsgNo) = 250 Then Exit Function
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Throw New Exception("Error downloading file")
            Return False
        End Try
    End Function

    Public Function BackDirectory() As Boolean
        Try
            ResetResponse()
            Send("CDUP")
            If Not (WaitForResponse.MsgNo) = 250 Then Exit Function
            ListDirectory()
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Throw New Exception("Error changing directory")
            Return False
        End Try
    End Function

#End Region

#Region "Handles server response"

    Private Function GetResponse() As String
        Try
            If MyFTPSocket.Connected Then
                Dim Response As New classStatus
                Dim WaitHandle() As WaitHandle
                Dim ReadResult As IAsyncResult = MyFTPSocket.BeginReceive(Response.Buffer, 0, Response.Buffer.Length, SocketFlags.None, AddressOf GetResponse, Response)
                Do : Loop Until Len(Response.ServerResponseMessage) > 0
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Sub GetResponse(ByVal ar As IAsyncResult)
        Try
            If MyFTPSocket.Connected Then
                Dim Socket As Socket = MyFTPSocket
                Dim Response As classStatus = CType(ar.AsyncState, classStatus)
                Dim BytesLeft As Integer = Socket.EndReceive(ar)

                If BytesLeft > 0 Then
                    Response.ServerResponseMessage = ASCII.GetString(Response.Buffer, 0, BytesLeft)
                    Socket.BeginReceive(Response.Buffer, 0, Response.Buffer.Length, SocketFlags.None, AddressOf GetResponse, Response)
                    ' Send the response back and do commands
                    bNewResponse = True
                    Dim MsgNo As Integer = CInt(Mid(Response.ServerResponseMessage, 1, 3))
                    Dim Msg As String = Response.ServerResponseMessage.Remove(0, 4)
                    ' Set all that need to be set
                    _EventLog.Append(Response.ServerResponseMessage)
                    LastResponse = New classLastResponse(MsgNo, Msg)
                    RaiseEvent OnEventAdded(MsgNo, Msg)
                    'Console.WriteLine(MsgNo & ":" & Msg)
                    SetCommand(MsgNo, Msg)
                    'RaiseEvent ReceivedData(MsgNo, Msg)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub SetCommand(ByVal MsgNo As Integer, ByVal Msg As String)
        Select Case MsgNo
            Case 230
                Loggedinto = True
            Case 227 ' PASV Mode
                PasvSocket = GetPasvSocket(Msg)
            Case 421 ' Server closing connection
                WaitForDisconnect()
        End Select
    End Sub

#End Region

#Region "Create Pasv Socket"

    Private Function GetPasvSocket(ByVal ServerResponseString As String) As Socket
        Try
            Dim ReceivedIP As String = ServerResponseString
            Dim IPasStrings() As String
            ' Strip the begining stuff of the string
            ReceivedIP = ReceivedIP.Remove(0, InStr(ReceivedIP, "(")) : ReceivedIP = ReceivedIP.Substring(0, ReceivedIP.IndexOf(")"))
            IPasStrings = Split(ReceivedIP, ",")
            Dim IP As String = IPasStrings(0) + "." + IPasStrings(1) + "." + IPasStrings(2) + "." + IPasStrings(3)
            Dim Port As Integer = CInt(IPasStrings(4)) * 256 + CInt(IPasStrings(5))
            Dim MyIP As New IPEndPoint(IPAddress.Parse(IP), Port)
            Dim MyPasvSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            MyPasvSocket.Connect(MyIP)
            Return MyPasvSocket
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

#End Region

#Region "Directories"

    Public Function ChangeDirectory(ByVal Directory As String) As Boolean
        Try
            ResetResponse()
            Send("CWD " & Directory)
            ' 550 is CWD is successfull
            If (WaitForResponse.MsgNo) = 550 Then
                ' Timeout happened or command is not successfull
                Return False
                RaiseEvent OnErrorMsg("Can't change directory - " & LastResponse.Msg)
                Throw New Exception("Can't change directory")
            Else
                ListDirectory()
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Public Function CurrentDirectory() As String
        Try
            ResetResponse()
            Send("PWD")
            ' 257 is PWD is successfull
            If (WaitForResponse.MsgNo) = 257 Then
                ' Timeout happened or command is not successfull
                Dim MyDir As String
                MyDir = LastResponse.Msg.Remove(0, InStr(LastResponse.Msg, """"))
                MyDir = MyDir.Substring(0, InStr(MyDir, """") - 1)
                Return MyDir
            Else
                Return ""
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function PhaseRegexDirectory(ByVal ReponseServerString As String, Optional ByVal TimeOut As Integer = 10000) As Boolean
        Try
            Dim Fileattb As classFileAttb
            Dim FileRegMatch As Match
            Dim FilelistString() As String = Split(ReponseServerString, vbCrLf)
            Dim i As Integer

            Dim Size As Integer
            Dim directory As Boolean = False
            Dim Timestamp As Date

            Directorys.Clear()
            Files.Clear()

            For i = 0 To UBound(FilelistString) - 1
                directory = False
                If regex_UNIX_1.Match(FilelistString(i)).Success Then
                    FileRegMatch = regex_UNIX_1.Match(FilelistString(i))
                ElseIf regex_UNIX_2.Match(FilelistString(i)).Success Then
                    FileRegMatch = regex_UNIX_2.Match(FilelistString(i))
                ElseIf regex_UNIX_3.Match(FilelistString(i)).Success Then
                    FileRegMatch = regex_UNIX_3.Match(FilelistString(i))
                ElseIf regex_WIND_1.Match(FilelistString(i)).Success Then
                    FileRegMatch = regex_WIND_1.Match(FilelistString(i))
                End If

                If FileRegMatch Is Nothing Then
                    Throw New Exception("Directory cast is not valid")
                    Exit Function
                End If

                If (Len(FileRegMatch.Groups("dir").Value) > 0) And (Not FileRegMatch.Groups("dir").Value = "-") Then directory = True
                If (FileRegMatch.Groups("size").Value = "") Then Size = 0 Else Size = CInt(FileRegMatch.Groups("size").Value)

                Fileattb = New classFileAttb(FileRegMatch.Groups("name").Value, Size, Timestamp, directory)

                If Len(Trim(UBound(FilelistString))) > 0 Then
                    If directory = True Then
                        Directorys.Add(Fileattb)
                    Else
                        Files.Add(Fileattb)
                    End If
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function ListDirectory(Optional ByVal TimeOut As Integer = 10000) As Boolean
        Try
            Dim BytesReceived As Integer
            Dim ReponseServerString As String
            Dim ReceiveBuffer(512) As Byte
            ResetResponse()
            ' Send command to server to enter passive mode
            Send("PASV")
            ' Check for successfull mode change
            If Not (WaitForResponseNo(227, 5000)) Then Exit Function
            ' Only open get listing when PASV port is open
            If Not (WaitForPasvSocket(5000) = True) Then Exit Function
            ' Send command to list
            ResetResponse()
            Send("LIST")
            WaitForResponse()

            Do
                BytesReceived = PasvSocket.Receive(ReceiveBuffer, 512, SocketFlags.None)
                ReponseServerString += Encoding.ASCII.GetString(ReceiveBuffer, 0, BytesReceived)
            Loop Until BytesReceived = 0

            PhaseRegexDirectory(ReponseServerString)

            If Not (WaitForResponseNo(226, 5000)) Then
                Throw New Exception("Error getting directory listing")
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

#End Region

#Region "Downloading/Uploading files"

#Region "File Checking exists etc"

    Public Function RemoteFileExist(ByVal Filename As String) As Boolean
        Dim MyFileAttb As classFileAttb
        For Each MyFileAttb In Files
            If UCase(MyFileAttb.Filename) = UCase(Filename) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function RemoteFileExistIndex(ByVal Filename As String) As Integer
        Dim MyFileAttb As classFileAttb
        For Each MyFileAttb In Files
            If UCase(MyFileAttb.Filename) = UCase(Filename) Then
                Return Files.IndexOf(MyFileAttb)
            End If
        Next
        Return -1
    End Function

#End Region

    Private Function GetFilename(ByVal Path As String) As String
        Dim Filename() As String = Path.Split("\")
        Return Filename(UBound(Filename))
    End Function

    Public Function UploadFile(ByVal Filename As String) As Boolean
        Try
            Dim BytesUploaded As Integer
            Dim BytesRemaining As Integer
            Dim FileSize As Integer
            Dim ReponseServerString As String
            Dim UploadBuffer(2048) As Byte

            If Not IO.File.Exists(Filename) Then
                Throw New Exception("File doesn't exist")
            End If

            'MyFileStream = New FileStream("fichero.pdf", FileMode.Open)

            If WaitForLogin(5000) = False Then Exit Function
            ' Reset the response watcher
            ResetResponse()
            ' Send command to server to enter passive mode
            Send("PASV")
            ' Check for successfull mode change
            If Not (WaitForResponseNo(227, 5000)) Then Exit Function
            ' Only open get listing when PASV port is open
            If Not (WaitForPasvSocket(5000) = True) Then Exit Function
            ' Send command to list
            Dim FileAttb As FileAttributes = IO.File.GetAttributes(Filename)

            Dim MyFile As New IO.FileStream(Filename, FileMode.Open)
            FileSize = MyFile.Length
            Dim Filename_short As String = GetFilename(Filename)
            Send("STOR " & Filename_short)
            If Not (WaitForResponseNo(125, 5000)) Then Exit Function
            Do
                BytesRemaining = MyFile.Read(UploadBuffer, 0, 512)
                BytesUploaded += BytesRemaining
                PasvSocket.Send(UploadBuffer, BytesRemaining, SocketFlags.None)
                RaiseEvent UploadProgress(Filename_short, BytesUploaded, FileSize)
                ApplicationDoEvents()
            Loop Until BytesRemaining = 0

            MyFile.Close()
            PasvSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            If Not (WaitForResponseNo(226, 5000)) Then
                Throw New Exception("Error uploading file")
                Return False
            Else
                If ChangeDirectory(CurrentDirectory) = True Then
                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Public Function DownloadFile(ByVal RemoteFilename As String, ByVal LocalFilename As String, Optional ByVal TimeOut As Integer = 10000) As Boolean
        Try
            Dim BytesReceived, BytesRemaining As Integer
            Dim ReponseServerString As String
            Dim ReceiveBuffer(2048) As Byte

            Dim FileIndex As Integer = RemoteFileExistIndex(RemoteFilename)

            If FileIndex = -1 Then
                Throw New Exception("File doesn't exist")
            End If

            ' Wait until loggedin
            If WaitForLogin(5000) = False Then Exit Function
            ' Reset the response watcher
            ResetResponse()
            ' Send command to server to enter passive mode
            Send("PASV")
            ' Check for successfull mode change
            If Not (WaitForResponseNo(227, 5000)) Then Exit Function
            ' Only open get listing when PASV port is open
            If Not (WaitForPasvSocket(5000) = True) Then Exit Function
            ' Send command to list
            Dim MyFile As New IO.FileStream(LocalFilename, FileMode.Create)
            Send("RETR " & RemoteFilename)
            If Not (WaitForResponseNo(125, 5000)) Then Exit Function
            Do
                BytesRemaining = PasvSocket.Receive(ReceiveBuffer, 512, SocketFlags.None)
                BytesReceived += BytesRemaining
                MyFile.Write(ReceiveBuffer, 0, BytesRemaining)
                RaiseEvent DownloadProgress(RemoteFilename, BytesReceived, DirectCast(Files(FileIndex), classFileAttb).Size)
                ApplicationDoEvents()
            Loop Until BytesReceived = DirectCast(Files(FileIndex), classFileAttb).Size ' BytesRemaining = 0

            MyFile.Close()

            If Not (WaitForResponseNo(226, 5000)) Then
                Throw New Exception("Error downloading file")
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

#End Region

#Region "Wait for .. functions"

    Private Sub ResetResponse()
        bNewResponse = False
    End Sub

    Private Function WaitForResponse(Optional ByVal TimeOut As Integer = 60000) As classLastResponse
        Try
            Dim StartTime As Long = Environment.TickCount
            While (bNewResponse = False) And (Not (Environment.TickCount - StartTime) = TimeOut)
                ApplicationDoEvents()
            End While
            ApplicationDoEvents()
            If (bNewResponse = False) Then
                ResetResponse()
                Return New classLastResponse(0, "")
            Else
                ResetResponse()
                Return LastResponse
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function WaitForResponseNo(ByVal MsgNo As Integer, Optional ByVal TimeOut As Integer = 60000) As Boolean
        Try
            Dim StartTime As Long = Environment.TickCount
            While (LastResponse.MsgNo <> MsgNo) And ((Environment.TickCount - StartTime) < TimeOut)
                ApplicationDoEvents()
            End While
            ApplicationDoEvents()
            If (LastResponse.MsgNo <> MsgNo) Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function WaitForPasvSocket(Optional ByVal TimeOut As Integer = 60000) As Boolean
        Try
            Dim StartTime As Long = Environment.TickCount
            While (PasvSocket Is Nothing) And ((Environment.TickCount - StartTime) < TimeOut)
                ApplicationDoEvents()
            End While
            If Not PasvSocket Is Nothing Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function WaitForLogin(Optional ByVal TimeOut As Integer = 60000) As Boolean
        Try
            Dim StartTime As Long = Environment.TickCount
            While (Len(_ServerType) = 0) And (Loggedinto = False) And ((Environment.TickCount - StartTime) < TimeOut)
                ApplicationDoEvents()
            End While
            If Loggedinto = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function WaitForDisconnect()
        StartsocketShutdown()
    End Function

#End Region

#Region "Gambiarra"

    Private Sub ApplicationDoEvents()
        Thread.Sleep(100)
        'Application.DoEvents()
        'Dim X As Object
        'X = New Object
        'X = vbNull
    End Sub
#End Region


End Class

