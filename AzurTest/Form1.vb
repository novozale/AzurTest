Imports System.Net.Http.Headers
Imports Microsoft.Graph
Imports Microsoft.Identity.Client
Imports Newtonsoft.Json

Public Class Form1
    Private Const client_id As String = "917b7b74-f5b2-45c5-ae05-6b4c895b4e79" '<-- enter the client_id guid here
    Private Const tenant_id As String = "f91cd4eb-0e4b-4bcc-982e-32c194cfcefa" '<-- enter either your tenant id here
    Private Const client_secret As String = "K3~MGDdZH553O61c-8_m.83S682tzUvg-2"
    Private authority As String = $"https://login.microsoftonline.com/f91cd4eb-0E4b-4bcc-982e-32c194cfcefa"

    Private _scopes As New List(Of String)
    Private ReadOnly Property scopes As List(Of String)
        Get
            If _scopes.Count = 0 Then
                _scopes.Add("User.read") '<-- add each scope you want to send as a seperate .add
                _scopes.Add("Calendars.ReadWrite")
            End If
            Return _scopes
        End Get
    End Property

    Private _pca As IPublicClientApplication = Nothing
    Private ReadOnly Property PCA As IPublicClientApplication
        Get
            If _pca Is Nothing Then
                _pca = PublicClientApplicationBuilder.Create(client_id).WithRedirectUri("msal917b7b74-f5b2-45c5-ae05-6b4c895b4e79://auth").WithAdfsAuthority(authority).Build()
            End If
            Return _pca
        End Get
    End Property

    Private _authProvider As InteractiveAuthenticationProvider = Nothing
    Private ReadOnly Property AuthProvider As InteractiveAuthenticationProvider
        Get
            If _authProvider Is Nothing Then
                _authProvider = New InteractiveAuthenticationProvider(PCA, scopes)
            End If
            Return _authProvider
        End Get
    End Property

    Private _graphClient As GraphServiceClient = Nothing
    Private ReadOnly Property GraphClient As GraphServiceClient
        Get
            If _graphClient Is Nothing Then
                _graphClient = New GraphServiceClient(AuthProvider)
            End If
            Return _graphClient
        End Get
    End Property

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetMe()

    End Sub

    Private Async Sub GetMe()
        Dim user As User

        'user = Await GraphClient.Me().Request().Select("displayName,employeeid").GetAsync()
        'RichTextBox1.Text = RichTextBox1.Text + ($"User = {user.DisplayName}, employeeid = {user.EmployeeId}")
        Dim events = Await GraphClient.Me.Calendar.Events.Request().GetAsync()
        'user = Await GraphClient.Me.Request().GetAsync()
        RichTextBox1.Text = RichTextBox1.Text + events.ToString()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        GetMe1()

    End Sub

    Private Async Sub GetMe1()
        Dim MyClient As GraphServiceClient

        MyClient = createMsGraphClient()
        Dim events = Await GraphClient.Me.Calendar.Events.Request().GetAsync()
        RichTextBox1.Text = RichTextBox1.Text + events.ToString()
    End Sub

    'Private Function createMsGraphClient() As GraphServiceClient

    '    Dim ConfidentialClientApplication As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
    '        .Create(client_id) _
    '        .WithTenantId(tenant_id) _
    '        .WithClientSecret(client_secret) _
    '    .Build()

    '    Dim graphServiceClient As GraphServiceClient = New GraphServiceClient(
    '        New DelegateAuthenticationProvider(
    '            Async Function(requestMessage)
    '                Dim authResult = Await ConfidentialClientApplication _
    '                .AcquireTokenForClient(scopes) _
    '                .ExecuteAsync()

    '                requestMessage.Headers.Authorization = New AuthenticationHeaderValue("bearer", authResult.AccessToken)
    '            End Function
    '        )
    '    )
    '    Return graphServiceClient
    'End Function

    Private Function createMsGraphClient() As GraphServiceClient

        Dim ConfidentialClientApplication As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
            .Create(client_id) _
            .WithTenantId(tenant_id) _
            .WithClientSecret(client_secret) _
        .Build()

        Dim graphServiceClient As GraphServiceClient = New GraphServiceClient(
            New DelegateAuthenticationProvider(
                Async Function(requestMessage) As Task
                    Dim authResult = Await ConfidentialClientApplication _
                    .AcquireTokenForClient(scopes) _
                    .ExecuteAsync()

                    requestMessage.Headers.Authorization = New AuthenticationHeaderValue("bearer", authResult.AccessToken)
                End Function
            )
        )
        Return graphServiceClient
    End Function

End Class
