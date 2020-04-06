Imports System.IO

Public Class Tela

    Private AlterandoTela As Boolean = True
    Private Ativou As Boolean = False
    Public WithEvents MyFTP As classFTP
    Private ArqEsc As FileInfo
    Private UltDt As Date = DateValue("01/01/2001")
    Private MeuIni As New Ini()
    Private Conectado As Boolean = False

    Private cLog As Log
    'Private CaminhoLog As String = ""
    'Private LogCriado As Boolean = False

    Private Sub Tela_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cLog = New Log()
        cLog.inicializa(False)

        MyFTP = New classFTP()
        cLog.Loga("Aciona FTP")
        'If MyFTP.Connect(Host, User, Pass) Then
        '    'Me.WindowState = FormWindowState.Minimized
        '    cLog.Loga("Conectou no FTP")
        '    Conectado = True
        'Else
        '    cLog.Loga("NÂO Conectou no FTP")
        '    Dim Erro As String = MyFTP.getErro()
        '    cLog.Loga("Erro = " & Erro)
        '    MsgBox("Deu Guru ... " & Host & " " & Erro & " ", MsgBoxStyle.Exclamation, "Error connecting")
        'End If
    End Sub

    Private Function Atualiza()
        Dim camLocal As String = MeuIni.getCamLocal()
        UltAtualizado(camLocal)
        Dim posApp As Integer = ArqEsc.DirectoryName.IndexOf("\public_html\")
        Dim parDir As String = ArqEsc.DirectoryName.Substring(posApp + 24).Replace("\", "/")
        Dim camFTP As String = MeuIni.getCamFTP()
        Dim DiretFTP As String = camFTP & parDir

        Label1.Text = ArqEsc.FullName
        If (ArqEsc.Length > 1000000) Then
            Me.Text = " Muito grande " & ArqEsc.Length.ToString() & " bytes"
            Atualiza = False
        Else
            Label1.Text = ArqEsc.FullName
            Directory.SetCurrentDirectory(ArqEsc.DirectoryName)
            MyFTP.ChangeDirectory(DiretFTP)
            Timer1.Enabled = True
            ProgressBar1.Maximum = ArqEsc.Length
            Me.Text = "Enviando " & ArqEsc.Name
            If MyFTP.UploadFile(ArqEsc.Name) = False Then
                MsgBox("Deu Guru")
                End
            End If
            Me.Text = ArqEsc.Name & " " & Format(Now, "HH:mm:ss")
            Timer1.Enabled = False
            Conectado = False
            MyFTP.Disconnect()
            Atualiza = True
        End If
    End Function

    Private Sub UploadProgress(ByVal Filename As String, ByVal BytesUploaded As Integer, ByVal Filesize As Integer) Handles MyFTP.UploadProgress
        ProgressBar1.Value = BytesUploaded
    End Sub

    Private Sub Erro(ByVal ErrorMsg As String) Handles MyFTP.OnErrorMsg
        MsgBox("Erro ... " & ErrorMsg, MsgBoxStyle.Exclamation, "FTPeitor")
    End Sub

    Private Function UltAtualizado(Pasta As String) As FileSystemInfo
        Dim dirInfo As DirectoryInfo = New DirectoryInfo(Pasta)
        BuscaArquivos(dirInfo)
    End Function

    Private Sub BuscaArquivos(diret As DirectoryInfo)

        Dim objDirectoryInfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(diret.FullName)
        'Dim objDirectoryInfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo("D:\Prog\Tele-Tudo\Site\app")

        SearchFiles(objDirectoryInfo)
        SearchDirectories(objDirectoryInfo)
    End Sub

    Private Function SearchFiles(ByVal objDirectoryInfo As System.IO.DirectoryInfo) As Boolean
        Dim objFileInfo As System.IO.FileInfo
        Dim bolReturn As Boolean = False
        For Each objFileInfo In objDirectoryInfo.GetFiles("*.php")
            Console.WriteLine(objFileInfo.Name)
            If objFileInfo.LastWriteTime > UltDt Then
                UltDt = objFileInfo.LastWriteTime
                ArqEsc = objFileInfo
            End If
        Next
        Return bolReturn
    End Function

    Private Function SearchDirectories(ByVal objDirectoryInfo As System.IO.DirectoryInfo) As Boolean
        Dim bolReturn As Boolean = False
        For Each objDirectoryInfo In objDirectoryInfo.GetDirectories
            bolReturn = True
            If objDirectoryInfo.Exists = True And objDirectoryInfo.Name <> "System Volume Information" And objDirectoryInfo.Name <> "RECYCLER" Then
                If SearchDirectories(objDirectoryInfo) = False And SearchFiles(objDirectoryInfo) = False Then
                    'CheckedListBoxDirectories.Items.Add(objDirectoryInfo.FullName, True)
                End If
            End If
        Next
        Return bolReturn
    End Function

    Private Sub Tela_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If AlterandoTela = False Then
            AlterandoTela = True
            If Conectado = False Then
                Conectar()
            End If
            If (Atualiza() = True) Then
                Me.WindowState = FormWindowState.Minimized
            End If
            ProgressBar1.Value = 0
            AlterandoTela = False
        End If
    End Sub

    Private Sub Conectar()
        If MyFTP.Connect(MeuIni.getHost(), MeuIni.getUsuario(), MeuIni.getSenha()) Then
            cLog.Loga("Conectou no FTP")
            Conectado = True
        Else
            cLog.Loga("NÂO Conectou no FTP")
            Dim Erro As String = MyFTP.getErro()
            cLog.Loga("Erro = " & Erro)
            MsgBox("Deu Guru ... " & Erro & " ", MsgBoxStyle.Exclamation, "Error connecting")
        End If
    End Sub

    Private Sub Tela_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If Ativou = False Then
            Ativou = True
            Me.Left = 852
            Me.Top = 748
            Me.WindowState = FormWindowState.Minimized
            AlterandoTela = False
        End If
    End Sub

    'Private Sub Loga(Texto As String)
    '    Dim Agora As String = DateTime.Now.ToLongTimeString()
    '    Console.WriteLine(Agora + " " + Texto)
    '    Dim aLog As StreamWriter
    '    Dim CaminhoLog As String = AppDomain.CurrentDomain.BaseDirectory.ToString() + "\log.log"
    '    If (LogCriado = True) Then
    '        aLog = File.AppendText(CaminhoLog)
    '    Else
    '        LogCriado = True
    '        aLog = File.CreateText(CaminhoLog)
    '    End If
    '    aLog.WriteLine(Agora + " " + Texto)
    '    aLog.Close()
    'End Sub

End Class