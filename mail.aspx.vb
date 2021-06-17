Imports MailKit.Net.Smtp
Imports MailKit
Imports MimeKit
Imports MailKit.Security
Imports System.Web.Configuration

Public Class mail
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mode As String = "", userid As String = "", pwd As String = "", server As String = "", port As Integer, ssl As Boolean
        mode = Request.QueryString("mode")

        Dim settings = System.Configuration.ConfigurationManager.AppSettings("mailSettings").ToString()
        Dim setting As String() = settings.Split(";")
        For i = 0 To setting.Length - 1
            If setting(i) <> "" Then
                If setting(i).Split("=")(0) = "userid" Then userid = setting(i).Split("=")(1)
                If setting(i).Split("=")(0) = "pwd" Then pwd = setting(i).Split("=")(1)
                If setting(i).Split("=")(0) = "server" Then server = setting(i).Split("=")(1)
                If setting(i).Split("=")(0) = "port" And setting(i).Split("=")(1) <> "" Then port = CInt(setting(i).Split("=")(1))
                If setting(i).Split("=")(0) = "ssl" And setting(i).Split("=")(1) <> "" Then ssl = CBool(setting(i).Split("=")(1))
            End If
        Next
        Dim msg As String = ""
        If mode <> "" Then
            If userid <> "" And pwd <> "" And server <> "" Then
                If mode = "imap" Then
                    msg = RetrieveIMAP(userid, pwd, server, port, ssl)
                    If msg = "" Then
                        Response.Write("{success:1}")
                    Else
                        Response.Write("{success:0, message:" & msg & "}")
                    End If
                End If
            Else
                msg = "Parameter is missing"
                Response.Write("{success:0, message:" & msg & "}")
            End If
        Else
            msg = "Mode is missing"
            Response.Write("{success:0, message:" & msg & "}")

        End If
    End Sub
    Public Function RetrieveIMAP(ByVal userid As String, ByVal pwd As String, ByVal server As String, ByVal port As Integer, ByVal ssl As Boolean) As String
        Dim r As String = ""
        Using client = New MailKit.Net.Imap.ImapClient()

            Try
                client.Connect(server, port, ssl)
                client.Authenticate(userid, pwd)
                Dim inbox = client.Inbox
                inbox.Open(FolderAccess.[ReadOnly])
                Console.WriteLine("Total messages: {0}", inbox.Count)
                Console.WriteLine("Recent messages: {0}", inbox.Recent)
                'setText("Total messages: " & inbox.Count)
                'setText("Recent messages: " & inbox.Recent)

                For i As Integer = 0 To inbox.Count - 1
                    Dim message = inbox.GetMessage(i)
                    Console.WriteLine("Subject: {0}", message.Subject)
                    'setText("Subject: " & message.Subject)
                Next

                client.Disconnect(True)
            Catch e As AuthenticationException
                Console.WriteLine(e.Message)
                r = e.Message
                'setText(e.Message)
            Catch e As NotSupportedException
                Console.WriteLine(e.Message)
                r = e.Message
                'setText(e.Message)
            End Try
        End Using
        Return r
    End Function
End Class