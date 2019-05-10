Public Class Form2
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Server As System.Windows.Forms.TextBox
    Friend WithEvents Username As System.Windows.Forms.TextBox
    Friend WithEvents Anonymous As System.Windows.Forms.CheckBox
    Friend WithEvents Password As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form2))
Me.Server = New System.Windows.Forms.TextBox
Me.Username = New System.Windows.Forms.TextBox
Me.Anonymous = New System.Windows.Forms.CheckBox
Me.Button1 = New System.Windows.Forms.Button
Me.Button2 = New System.Windows.Forms.Button
Me.Label1 = New System.Windows.Forms.Label
Me.Label2 = New System.Windows.Forms.Label
Me.Label3 = New System.Windows.Forms.Label
Me.Password = New System.Windows.Forms.TextBox
Me.SuspendLayout
'
'Server
'
Me.Server.Location = New System.Drawing.Point(96, 8)
Me.Server.Name = "Server"
Me.Server.Size = New System.Drawing.Size(272, 20)
Me.Server.TabIndex = 0
Me.Server.Text = "ftp.microsoft.com"
'
'Username
'
Me.Username.Location = New System.Drawing.Point(96, 40)
Me.Username.Name = "Username"
Me.Username.ReadOnly = true
Me.Username.Size = New System.Drawing.Size(160, 20)
Me.Username.TabIndex = 1
Me.Username.Text = "anonymous"
'
'Anonymous
'
Me.Anonymous.Checked = true
Me.Anonymous.CheckState = System.Windows.Forms.CheckState.Checked
Me.Anonymous.Location = New System.Drawing.Point(264, 42)
Me.Anonymous.Name = "Anonymous"
Me.Anonymous.Size = New System.Drawing.Size(104, 16)
Me.Anonymous.TabIndex = 3
Me.Anonymous.Text = "Anonymous"
'
'Button1
'
Me.Button1.Location = New System.Drawing.Point(16, 104)
Me.Button1.Name = "Button1"
Me.Button1.Size = New System.Drawing.Size(128, 23)
Me.Button1.TabIndex = 4
Me.Button1.Text = "Connect"
'
'Button2
'
Me.Button2.Location = New System.Drawing.Point(232, 104)
Me.Button2.Name = "Button2"
Me.Button2.Size = New System.Drawing.Size(128, 23)
Me.Button2.TabIndex = 5
Me.Button2.Text = "Cancel"
'
'Label1
'
Me.Label1.Location = New System.Drawing.Point(16, 10)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(72, 16)
Me.Label1.TabIndex = 6
Me.Label1.Text = "FTP Server :"
'
'Label2
'
Me.Label2.Location = New System.Drawing.Point(16, 42)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(72, 16)
Me.Label2.TabIndex = 7
Me.Label2.Text = "Username :"
'
'Label3
'
Me.Label3.Location = New System.Drawing.Point(16, 66)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(72, 16)
Me.Label3.TabIndex = 8
Me.Label3.Text = "Password :"
'
'Password
'
Me.Password.Location = New System.Drawing.Point(96, 64)
Me.Password.Name = "Password"
Me.Password.PasswordChar = Microsoft.VisualBasic.ChrW(42)
Me.Password.ReadOnly = true
Me.Password.Size = New System.Drawing.Size(160, 20)
Me.Password.TabIndex = 2
Me.Password.Text = "test@test.com"
'
'Form2
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(376, 141)
Me.Controls.Add(Me.Label3)
Me.Controls.Add(Me.Label2)
Me.Controls.Add(Me.Label1)
Me.Controls.Add(Me.Button2)
Me.Controls.Add(Me.Button1)
Me.Controls.Add(Me.Anonymous)
Me.Controls.Add(Me.Password)
Me.Controls.Add(Me.Username)
Me.Controls.Add(Me.Server)
Me.Icon = CType(resources.GetObject("$this.Icon"),System.Drawing.Icon)
Me.Name = "Form2"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Connect to FTP Server"
Me.ResumeLayout(false)

    End Sub

#End Region

    Public MyFTP As FTP2.classFTP

    Private Sub EnableDisplay()
        Server.Enabled = True
        Username.Enabled = True
        Password.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
    End Sub
    Private Sub DisableDisplay()
        Server.Enabled = False
        Username.Enabled = False
        Password.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If MyFTP.Connected = True Then MyFTP.Disconnect()
        If Len(Trim(Server.Text)) > 0 And _
           Len(Trim(Username.Text)) > 0 Then
            DisableDisplay()
            If MyFTP.Connect(Server.Text, Username.Text, Password.Text) Then
                Me.Close()
            Else
                MsgBox("Problem connecting to ... " & Server.Text, MsgBoxStyle.Exclamation, "Error connecting")
                EnableDisplay()
            End If
        End If
    End Sub

    Private Sub Anonymous_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Anonymous.CheckedChanged
        If Anonymous.Checked = True Then
            Username.Text = "anonymous"
            Username.ReadOnly = True
            Password.Text = "test@test.com"
            Password.ReadOnly = True
        Else
            Username.ReadOnly = False
            Password.ReadOnly = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
