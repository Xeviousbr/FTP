Imports System.IO

Public Class Tela

    Private AlterandoTela As Boolean = True
    Private Ativou As Boolean = False
    Public WithEvents MyFTP As FTP2.classFTP
    Private Host As String = "ftp.intonses.com.br"
    Private User As String = "inton634"
    Private Pass As String = "4zk3xkV3K5"
    Private ArqEsc As System.IO.FileInfo
    Private UltDt As Date = DateValue("01/01/2001")

    Private Sub Tela_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyFTP = New classFTP()
        If MyFTP.Connect(Host, User, Pass) Then
            'Me.WindowState = FormWindowState.Minimized
        Else
            MsgBox("Deu Guru ... " & Host, MsgBoxStyle.Exclamation, "Error connecting")
        End If
    End Sub

    Private Sub Atualiza()
        UltAtualizado("D:\Prog\Tele-Tudo\Site\app")
        Label1.Text = ArqEsc.FullName
        Directory.SetCurrentDirectory(ArqEsc.DirectoryName)
        Dim posApp As Integer = ArqEsc.DirectoryName.IndexOf("\app\")
        Dim parDir As String = ArqEsc.DirectoryName.Substring(posApp + 5).Replace("\", "/")
        Dim DiretFTP As String = "public_html/teletudo/app/" & parDir
        MyFTP.ChangeDirectory(DiretFTP)
        Timer1.Enabled = True
        ProgressBar1.Maximum = ArqEsc.Length
        If MyFTP.UploadFile(ArqEsc.Name) = False Then
            MsgBox("Deu Guru")
            End
        End If
        Timer1.Enabled = False
    End Sub

    Private Sub UploadProgress(ByVal Filename As String, ByVal BytesUploaded As Integer, ByVal Filesize As Integer) Handles MyFTP.UploadProgress
        ProgressBar1.Value = BytesUploaded
    End Sub

    Private Function UltAtualizado(Pasta As String) As FileSystemInfo
        Dim dirInfo As DirectoryInfo = New DirectoryInfo(Pasta)
        BuscaArquivos(dirInfo)
    End Function

    Private Sub BuscaArquivos(diret As DirectoryInfo)
        Dim objDirectoryInfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo("D:\Prog\Tele-Tudo\Site\app")
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
            Atualiza()
            Me.WindowState = FormWindowState.Minimized
            ProgressBar1.Value = 0
            AlterandoTela = False
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
End Class