Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents FilesFoldersImagelist As System.Windows.Forms.ImageList
    Friend WithEvents RemoteListview As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ToolBar1 As System.Windows.Forms.ToolBar
    Friend WithEvents LogWindow As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ClientListview As System.Windows.Forms.ListView
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents PutButton As System.Windows.Forms.Button
    Friend WithEvents GetButton As System.Windows.Forms.Button
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.LogWindow = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ToolBar1 = New System.Windows.Forms.ToolBar
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.RemoteListview = New System.Windows.Forms.ListView
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.FilesFoldersImagelist = New System.Windows.Forms.ImageList(Me.components)
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ClientListview = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.GetButton = New System.Windows.Forms.Button
        Me.PutButton = New System.Windows.Forms.Button
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.Panel1.SuspendLayout()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LogWindow
        '
        Me.LogWindow.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LogWindow.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LogWindow.Location = New System.Drawing.Point(0, 28)
        Me.LogWindow.Multiline = True
        Me.LogWindow.Name = "LogWindow"
        Me.LogWindow.ReadOnly = True
        Me.LogWindow.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.LogWindow.Size = New System.Drawing.Size(680, 100)
        Me.LogWindow.TabIndex = 1
        Me.LogWindow.Text = ""
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.LogWindow)
        Me.Panel1.Controls.Add(Me.ToolBar1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(680, 128)
        Me.Panel1.TabIndex = 5
        '
        'ToolBar1
        '
        Me.ToolBar1.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
        Me.ToolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton1})
        Me.ToolBar1.DropDownArrows = True
        Me.ToolBar1.ImageList = Me.ImageList1
        Me.ToolBar1.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar1.Name = "ToolBar1"
        Me.ToolBar1.ShowToolTips = True
        Me.ToolBar1.Size = New System.Drawing.Size(680, 28)
        Me.ToolBar1.TabIndex = 0
        Me.ToolBar1.TextAlign = System.Windows.Forms.ToolBarTextAlign.Right
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.ImageIndex = 0
        Me.ToolBarButton1.Text = "Connect to FTP Server"
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 439)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(680, 22)
        Me.StatusBar1.TabIndex = 6
        Me.StatusBar1.Text = "StatusBar1"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(206, 443)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(453, 16)
        Me.ProgressBar1.TabIndex = 7
        '
        'RemoteListview
        '
        Me.RemoteListview.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RemoteListview.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3})
        Me.RemoteListview.HideSelection = False
        Me.RemoteListview.Location = New System.Drawing.Point(0, 128)
        Me.RemoteListview.Name = "RemoteListview"
        Me.RemoteListview.Size = New System.Drawing.Size(304, 312)
        Me.RemoteListview.SmallImageList = Me.FilesFoldersImagelist
        Me.RemoteListview.TabIndex = 8
        Me.RemoteListview.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Directories / Files on Remote Server"
        Me.ColumnHeader2.Width = 200
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "File Size"
        Me.ColumnHeader3.Width = 100
        '
        'FilesFoldersImagelist
        '
        Me.FilesFoldersImagelist.ImageSize = New System.Drawing.Size(16, 16)
        Me.FilesFoldersImagelist.ImageStream = CType(resources.GetObject("FilesFoldersImagelist.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.FilesFoldersImagelist.TransparentColor = System.Drawing.Color.Transparent
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Directories / Files on Local PC"
        Me.ColumnHeader1.Width = 200
        '
        'ClientListview
        '
        Me.ClientListview.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClientListview.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader4})
        Me.ClientListview.HideSelection = False
        Me.ClientListview.Location = New System.Drawing.Point(360, 128)
        Me.ClientListview.Name = "ClientListview"
        Me.ClientListview.Size = New System.Drawing.Size(320, 312)
        Me.ClientListview.SmallImageList = Me.FilesFoldersImagelist
        Me.ClientListview.TabIndex = 9
        Me.ClientListview.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "File Size"
        Me.ColumnHeader4.Width = 116
        '
        'GetButton
        '
        Me.GetButton.Location = New System.Drawing.Point(312, 256)
        Me.GetButton.Name = "GetButton"
        Me.GetButton.Size = New System.Drawing.Size(40, 23)
        Me.GetButton.TabIndex = 10
        Me.GetButton.Text = "GET"
        '
        'PutButton
        '
        Me.PutButton.Location = New System.Drawing.Point(312, 288)
        Me.PutButton.Name = "PutButton"
        Me.PutButton.Size = New System.Drawing.Size(40, 23)
        Me.PutButton.TabIndex = 11
        Me.PutButton.Text = "PUT"
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.Width = 200
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarPanel2.Width = 464
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(680, 461)
        Me.Controls.Add(Me.PutButton)
        Me.Controls.Add(Me.GetButton)
        Me.Controls.Add(Me.ClientListview)
        Me.Controls.Add(Me.RemoteListview)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FTP Demo Client"
        Me.Panel1.ResumeLayout(False)
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private WithEvents MyFTP As New classFTP
    Private StartTime As Long = 0
    Private Transfer As Boolean = False

    Private Sub MySub()
        Dim MyFTP As New classFTP
        If MyFTP.Connect("ftp.microsoft.com", "anonymouse", "test@test.com", "reskit/win2000") Then
            MyFTP.DownloadFile("now.zip", "now.zip")
        Else
            MsgBox("Could not connect", MsgBoxStyle.Information, "Error Connecting")
        End If
    End Sub

#Region "Events"

    Private Sub MyFTP_ErrorMsg(ByVal ErrorMsg As String) Handles MyFTP.OnErrorMsg
        MsgBox(ErrorMsg)
    End Sub

    Private Sub MyFTP_DownloadProgress(ByVal Filename As String, ByVal BytesReveived As Integer, ByVal Filesize As Integer) Handles MyFTP.DownloadProgress
        If BytesReveived = Filesize Then
            ProgressBar1.Value = 0
            Me.Text = "FTP Demo Client"
        Else
            Me.Text = "FTP Demo Client - Downloading .. " & Filename & " (" & Math.Round((BytesReveived / 1024)) & "kb received) .. " & Math.Round((BytesReveived / Filesize * 100)) & "%"
            ProgressBar1.Maximum = Filesize
            ProgressBar1.Value = BytesReveived
        End If
    End Sub

    Private Sub MyFTP_UploadProgress(ByVal Filename As String, ByVal BytesUploaded As Integer, ByVal Filesize As Integer) Handles MyFTP.UploadProgress
        If BytesUploaded = Filesize Then
            ProgressBar1.Value = 0
            Me.Text = "FTP Demo Client"
        Else
            Me.Text = "FTP Demo Client - Uploading .. " & Filename & " (" & Math.Round((BytesUploaded / 1024)) & "kb received) .. " & Math.Round((BytesUploaded / Filesize * 100)) & "%"
            ProgressBar1.Maximum = Filesize
            ProgressBar1.Value = BytesUploaded
        End If
    End Sub

    Private Sub MyFTP_OnEventAdded(ByVal MsgNo As Integer, ByVal Msg As String) Handles MyFTP.OnEventAdded
        LogWindow.Text += MsgNo & " " & Msg
        LogWindow.Select(Len(LogWindow.Text), 1)
        LogWindow.ScrollToCaret()
    End Sub

    Private Sub MyFTP_OnDisconnected() Handles MyFTP.OnDisconnected
        DisableDisplay()
        MsgBox("You are disconnected from the server", MsgBoxStyle.Information, "Disconnected")
    End Sub

#End Region

#Region "Remote updates"

    Private Sub FTPRefresh()
        Try
            Dim ItmX As ListViewItem
            RemoteListview.Items.Clear()
            If MyFTP.Connected Then
                ItmX = New ListViewItem(" .. ", 0)
                ItmX.Tag = "CDUP"
                RemoteListview.Items.Add(ItmX)
                Dim DirFile As classFTP.classFileAttb
                For Each DirFile In MyFTP.Directorys
                    ItmX = New ListViewItem(DirFile.Filename, 0)
                    ItmX.Tag = "DIR"
                    ItmX.SubItems.Add(DirFile.Size)
                    RemoteListview.Items.Add(ItmX)
                Next
                For Each DirFile In MyFTP.Files
                    ItmX = New ListViewItem(DirFile.Filename, -1)
                    ItmX.Tag = "FILE"
                    ItmX.SubItems.Add(Math.Round(DirFile.Size / 1024) & " kb")
                    RemoteListview.Items.Add(ItmX)
                Next
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub RemoteListview_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RemoteListview.DoubleClick
        If RemoteListview.SelectedItems.Count > 0 Then
            Select Case RemoteListview.SelectedItems(0).Tag
                Case "CDUP"
                    MyFTP.BackDirectory()
                    FTPRefresh()
                Case "FILE"
                    If Transfer = False Then
                        Transfer = True
                        StartTime = Environment.TickCount
                        MyFTP.DownloadFile(RemoteListview.SelectedItems(0).Text, Path & RemoteListview.SelectedItems(0).Text)
                        ClientRefresh()
                        Transfer = False
                    Else
                        MsgBox("Tranfer already in progress please wait...", MsgBoxStyle.Information, "Please wait")
                    End If
                Case "DIR"
                    MyFTP.ChangeDirectory(RemoteListview.SelectedItems(0).Text)
                    FTPRefresh()
            End Select
        End If
    End Sub

#End Region

#Region "Client updates"

    Private Path As String = "C:\"

    Private Sub PathUP()
        Dim MyPath() As String = Split(Path, "\")
        If UBound(MyPath) > 1 Then
            ReDim Preserve MyPath(UBound(MyPath) - 2)
            Path = Join(MyPath, "\")
        End If
    End Sub

    Private Sub ClientChangeDirectory(ByVal Directory As String)
        Path = Path & "\" & Directory & "\"
    End Sub

    Private Sub ClientRefresh()
        Try
            Dim Directory As String
            Dim ItmX As ListViewItem

            ClientListview.Items.Clear()

            ItmX = New ListViewItem(" .. ", 0)
            ItmX.Tag = "CDUP"
            ClientListview.Items.Add(ItmX)

            For Each Directory In IO.Directory.GetDirectories(Path)
                ItmX = New ListViewItem(Replace(Directory, Path, ""), 0)
                ItmX.Tag = "DIR"
                ItmX.SubItems.Add("0")
                ClientListview.Items.Add(ItmX)
            Next
            Dim File As String
            For Each File In IO.Directory.GetFiles(Path)
                ItmX = New ListViewItem(Replace(File, Path, ""), -1)
                ItmX.Tag = "FILE"
                Dim FileAttb As New IO.FileInfo(File)
                ItmX.SubItems.Add(Math.Round(FileAttb.Length / 1024) & " kb")
                ClientListview.Items.Add(ItmX)
            Next
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub ClientListview_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClientListview.DoubleClick
        If ClientListview.SelectedItems.Count > 0 Then
            Select Case ClientListview.SelectedItems(0).Tag
                Case "CDUP"
                    PathUP()
                    ClientRefresh()
                Case "FILE"
                    If Transfer = False Then
                        Transfer = True
                        MyFTP.UploadFile(Path & ClientListview.SelectedItems(0).Text)
                        FTPRefresh()
                        Transfer = False
                    Else
                        MsgBox("Tranfer already in progress please wait...", MsgBoxStyle.Information, "Please wait")
                    End If
                Case "DIR"
                    ClientChangeDirectory(ClientListview.SelectedItems(0).Text)
                    ClientRefresh()
            End Select
        End If
    End Sub

#End Region

#Region "Display"

    Private Sub DisableDisplay()
        RemoteListview.Enabled = False
        ClientListview.Enabled = False
        RemoteListview.Items.Clear()
        ClientListview.Items.Clear()
    End Sub

    Private Sub EnableDisplay()
        RemoteListview.Enabled = True
        ClientListview.Enabled = True
        ClientRefresh()
    End Sub

#End Region

#Region "Form load"

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DisableDisplay()
    End Sub

#End Region

#Region "Toolbar / Start Connection"

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick
        Select Case e.Button.Text
            Case "Connect to FTP Server"
                Dim FTPConnectFrm As New Form2
                FTPConnectFrm.MyFTP = MyFTP
                FTPConnectFrm.ShowDialog()
                If MyFTP.Connected Then
                    EnableDisplay()
                    FTPRefresh()
                End If
        End Select
    End Sub

#End Region

#Region "Get Put Buttons"

    Private Sub GetButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetButton.Click
        If Transfer = False Then
            If RemoteListview.SelectedItems.Count > 0 Then
                Dim ItmX As ListViewItem
                Dim i As Integer = 1
                Transfer = True
                RemoteListview.Enabled = False
                For Each ItmX In RemoteListview.SelectedItems
                    StatusBar1.Panels(0).Text = "Downloading file (" & i & "/" & RemoteListview.SelectedItems.Count & ")" & vbCrLf
                    MyFTP.DownloadFile(ItmX.Text, Path & ItmX.Text)
                    i += 1
                Next
                ClientRefresh()
                RemoteListview.Enabled = True
                Transfer = False
                StatusBar1.Panels(0).Text = ""
            End If
        Else
            MsgBox("Tranfer already in progress please wait...", MsgBoxStyle.Information, "Please wait")
        End If
    End Sub

    Private Sub PutButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PutButton.Click
        If Transfer = False Then
            If ClientListview.SelectedItems.Count > 0 Then
                Dim i As Integer = 1
                Dim ItmX As ListViewItem
                ClientListview.Enabled = False
                Transfer = True
                For Each ItmX In ClientListview.SelectedItems
                    StatusBar1.Panels(0).Text = "Uploading file (" & i & "/" & ClientListview.SelectedItems.Count & ")" & vbCrLf
                    MyFTP.UploadFile(Path & ItmX.Text)
                    i += 1
                Next
                FTPRefresh()
                ClientListview.Enabled = True
                Transfer = False
                StatusBar1.Panels(0).Text = ""
            End If
        Else
            MsgBox("Tranfer already in progress please wait...", MsgBoxStyle.Information, "Please wait")
        End If
    End Sub

#End Region

End Class

