Imports MailKit.Net.Smtp
Imports MailKit
Imports MimeKit
Imports MailKit.Security
Imports System.Web.Configuration
Imports Microsoft.Identity.Client

Public Class mail
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mode As String = "", userid As String = "", pwd As String = "", server As String = "", port As Integer, ssl As Boolean
        Dim clientId As String, dirId As String
        mode = Request.QueryString("mode")

        clientId = System.Configuration.ConfigurationManager.AppSettings("clientId").ToString()
        dirId = System.Configuration.ConfigurationManager.AppSettings("directoryId").ToString()
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
        If Not (mode Is Nothing) And mode <> "" Then
            If userid <> "" And pwd <> "" And server <> "" Then
                If mode = "imap" Then

                    msg = RetrieveIMAPAsync(userid, pwd, server, port, ssl, clientId, dirId).GetAwaiter().GetResult()
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
    Public Async Function RetrieveIMAPAsync(ByVal userid As String, ByVal pwd As String, ByVal server As String, ByVal port As Integer, ByVal ssl As Boolean, clientId As String, dirId As String) As Threading.Tasks.Task(Of String)
        Dim r As String = ""

        Dim options = New PublicClientApplicationOptions With {
            .ClientId = clientId,
            .TenantId = dirId,
            .RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        }
        Dim publicClientApplication = PublicClientApplicationBuilder.CreateWithApplicationOptions(options).Build()
        Dim scopes = New String() {"email", "offline_access", "https://outlook.office.com/IMAP.AccessAsUser.All"}
        Dim authToken = Await publicClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync()
        Dim oauth2 = New SaslMechanismOAuth2(authToken.Account.Username, authToken.AccessToken)

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