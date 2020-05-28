'TANIPChecker
'Todd Nelson, Nelonic Systems, LLC
'http://nelonic.com

Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Mail

Module Module1

    Dim result As String = ""
    Dim fromEmail As String = ""
    Dim toEmail As String = ""
    Dim server As String = ""
    Dim port As String = "25"
    Dim username As String = ""
    Dim password As String = ""
    Dim useSSL As String = "false"
    Dim smtp As New SmtpClient()

    Sub Main(ByVal args() As String)

        For Each arg As String In args
            If arg.Substring(0, 1) = "/" And arg.Contains("=") Then
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/fromemail=" And arg.Length > 11 Then fromEmail = arg.Substring(arg.IndexOf("=") + 1)
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/toemail=" And arg.Length > 9 Then toEmail = arg.Substring(arg.IndexOf("=") + 1)
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/server=" And arg.Length > 8 Then server = arg.Substring(arg.IndexOf("=") + 1)
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/port=" And arg.Length > 6 Then port = arg.Substring(arg.IndexOf("=") + 1)
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/username=" And arg.Length > 10 Then username = arg.Substring(arg.IndexOf("=") + 1)
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/password=" And arg.Length > 10 Then password = arg.Substring(arg.IndexOf("=") + 1)
                If arg.Substring(0, arg.IndexOf("=") + 1) = "/usessl=" And arg.Length > 8 Then useSSL = arg.Substring(arg.IndexOf("=") + 1)

            End If

            If arg = "/help" Then
                Console.WriteLine(vbCrLf & "TAN IP Checker, Ver. 1.00 (8/12/2016), Nelonic Systems, LLC.  http://nelonic.com" & vbCrLf)
                Console.WriteLine("Syntax:")
                Console.WriteLine(vbTab & "/fromemail=<email address to send from> (Required)")
                Console.WriteLine(vbTab & "/toemail=<email address to send to> (Required)")
                Console.WriteLine(vbTab & "/server=<address of SMTP server> (Required)")
                Console.WriteLine(vbTab & "/port=<SMTP port> (Optional, default: 25)")
                Console.WriteLine(vbTab & "/username=<SMTP username> (Optional)")
                Console.WriteLine(vbTab & "/password=<SMTP password> (Optional)")
                Console.WriteLine(vbTab & "/usessl=<true/false> (Optional, default: false, enables SMTP SSL/TLS)")
                Exit Sub
            End If
        Next

        If fromEmail = "" Or toEmail = "" Or server = "" Then
            Console.WriteLine("Email addresses and server address are required.")
            Exit Sub
        End If

        Dim currentIP As String = ""
        Dim msg As String = ""

        smtp.Port = port
        smtp.Host = server
        smtp.UseDefaultCredentials = False
        smtp.EnableSsl = CBool(useSSL)

        If username <> "" And password <> "" Then smtp.Credentials = New Net.NetworkCredential(username, password)

        If Not File.Exists("CurrentIP.txt") Then
            Console.WriteLine("First run, saving current IP to file.")

            GetIP()

            Dim sw As New StreamWriter("CurrentIP.txt")
            sw.Write(result)
            sw.Close()
            Exit Sub
        Else
            Dim sr As New StreamReader("CurrentIP.txt")
            currentIP = sr.ReadToEnd.Trim
            sr.Close()
        End If


        GetIP()

        'If Not IPAddress.TryParse(currentIP, Nothing) Then
        '    Console.WriteLine("Exception: IP from stored file not an IP address: " & currentIP)
        '    System.Diagnostics.EventLog.WriteEntry("TAN IP Checker", "Exception: IP from stored file not an IP address." & vbCrLf & vbCrLf & currentIP, EventLogEntryType.Error, 1, 1)
        '    smtp.Send(fromEmail, toEmail, "IP Changed ERROR", "Exception: IP from stored file not an IP address." & vbCrLf & vbCrLf & currentIP)
        '    Exit Sub
        'End If

        If Not IPAddress.TryParse(result, Nothing) Then
            Console.WriteLine("Exception: IP from internet not an IP address: " & result)
            System.Diagnostics.EventLog.WriteEntry("TAN IP Checker", "Exception: IP from internet not an IP address." & vbCrLf & vbCrLf & result, EventLogEntryType.Error, 1, 1)
            SendEmail("IP Changed ERROR", "Exception: IP from internet not an IP address." & vbCrLf & vbCrLf & result)
            Exit Sub
        End If

        If currentIP <> result Then
            msg = "Old IP: " & currentIP & vbCrLf & "New IP: " & result

            Dim sw As New StreamWriter("CurrentIP.txt")
            sw.Write(result)
            sw.Close()

            System.Diagnostics.EventLog.WriteEntry("TAN IP Checker", "IP Changed" & vbCrLf & vbCrLf & msg, EventLogEntryType.Warning, 1, 1)
            SendEmail("IP Changed", msg)
        Else
            Console.WriteLine("IP hasn't changed, still " & currentIP)
        End If

        Console.WriteLine(msg)

    End Sub

    Sub SendEmail(subject As String, body As String)

        Try
            smtp.Send(fromEmail, toEmail, subject, body)
        Catch e As Exception
            Console.WriteLine(e.Message & vbCrLf & vbCrLf & e.StackTrace)
        End Try

    End Sub

    Sub GetIP()

        Dim tryCount As Integer = 0

        Try

10:
            tryCount += 1
            Dim webClient As New System.Net.WebClient

            result = webClient.DownloadString("https://api.ipify.org")

            result = result.Trim

        Catch e As Exception

            If tryCount < 3 Then
                Threading.Thread.Sleep(1000)
                GoTo 10
            End If

            Console.WriteLine(e.Message & vbCrLf & vbCrLf & e.StackTrace)

            System.Diagnostics.EventLog.WriteEntry("TAN IP Checker", e.Message & vbCrLf & vbCrLf & e.StackTrace, EventLogEntryType.Error, 1, 1)

            SendEmail("IP Changed ERROR", e.Message & vbCrLf & vbCrLf & e.StackTrace)
        End Try
    End Sub

End Module

