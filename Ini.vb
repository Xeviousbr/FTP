Imports System.Text

Public Class Ini

    Private Declare Auto Function GetPrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    'Private Declare Auto Function WritePrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private nmArquivo As String = AppDomain.CurrentDomain.BaseDirectory() & "Ftepeia.ini"

    Private Function LeArquivoINI(ByVal section_name As String, ByVal key_name As String, ByVal default_value As String) As String
        Const MAX_LENGTH As Integer = 500
        Dim string_builder As New StringBuilder(MAX_LENGTH)
        GetPrivateProfileString(section_name, key_name, default_value, string_builder, MAX_LENGTH, nmArquivo)
        Return string_builder.ToString()
    End Function

    Public Function getUsuario() As String
        Return LeArquivoINI("FTP", "User", "teletu76")
    End Function

    Public Function getSenha() As String
        Return LeArquivoINI("FTP", "Senha", "bPruC717y0")
    End Function

    Public Function getHost() As String
        Return LeArquivoINI("FTP", "Host", "ftp.tele-tudo.com")
    End Function

    Public Function getCamFTP() As String
        Return LeArquivoINI("Config", "CamFTP", "public_html/teletudo/app/")
    End Function

    Public Function getCamLocal() As String
        Return LeArquivoINI("Config", "CamLocal", "D:\Prog\Tele-Tudo\Site\app")
    End Function

End Class
