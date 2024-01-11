Imports System.Net.Http
Imports Microsoft.Graph
Imports Microsoft.Identity.Client

Public Class InteractiveAuthenticationProvider
    Implements IAuthenticationProvider

    Private Property Pca As IPublicClientApplication
    Private Property Scopes As List(Of String)

    Private Sub New()
        'Intentionally left blank to prevent empty constructor
    End Sub

    Public Sub New(pca As IPublicClientApplication, scopes As List(Of String))
        Me.Pca = pca
        Me.Scopes = scopes
    End Sub

    Public Async Function AuthenticateRequestAsync(request As HttpRequestMessage) As Task Implements IAuthenticationProvider.AuthenticateRequestAsync
        Dim accounts As IEnumerable(Of IAccount)
        Dim result As AuthenticationResult = Nothing

        accounts = Await Pca.GetAccountsAsync()
        Dim interactionRequired As Boolean = False

        Try
            result = Await Pca.AcquireTokenSilent(Scopes, accounts.FirstOrDefault).ExecuteAsync()
        Catch ex1 As MsalUiRequiredException
            interactionRequired = True
        Catch ex2 As Exception
            MsgBox($"Authentication error: {ex2.Message}", MsgBoxStyle.OkOnly, "Attention!")
        End Try

        If interactionRequired Then
            Try
                result = Await Pca.AcquireTokenInteractive(Scopes).ExecuteAsync()
            Catch ex As Exception
                MsgBox($"Authentication error: {ex.Message}", MsgBoxStyle.OkOnly, "Attention!")
            End Try
        End If

        Form1.RichTextBox1.Text = Form1.RichTextBox1.Text + ($"Access Token: {result.AccessToken}{Environment.NewLine}")
        Form1.RichTextBox1.Text = Form1.RichTextBox1.Text + ($"Graph Request: {request.RequestUri}")
        'You must set the access token for the authorization of the current request
        request.Headers.Authorization = New Headers.AuthenticationHeaderValue("Bearer", result.AccessToken)
    End Function
End Class
