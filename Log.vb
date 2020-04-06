Imports System.IO

Public Class Log

    Private CaminhoLog As String = ""
    Private LogCriado As Boolean = False

    Public Sub inicializa(pLogCriado As Boolean)
        LogCriado = pLogCriado
    End Sub

    Public Sub Loga(Texto As String)
        Dim Agora As String = DateTime.Now.ToLongTimeString()
        Console.WriteLine(Agora + " " + Texto)
        Dim aLog As StreamWriter
        Dim CaminhoLog As String = AppDomain.CurrentDomain.BaseDirectory.ToString() + "\log.log"
        If (LogCriado = True) Then
            aLog = File.AppendText(CaminhoLog)
        Else
            LogCriado = True
            aLog = File.CreateText(CaminhoLog)
        End If
        aLog.WriteLine(Agora + " " + Texto)
        aLog.Close()
    End Sub

End Class
