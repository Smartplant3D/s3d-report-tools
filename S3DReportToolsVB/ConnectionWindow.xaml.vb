Imports S3DReportToolsVB.Models
Imports S3DReportToolsVB.Services

Public Class ConnectionWindow
    Public Property UpdatedConfig As DatabaseConfig
    Private ReadOnly _dbService As New DatabaseService()

    Public Sub New(config As DatabaseConfig)
        InitializeComponent()
        UpdatedConfig = config
        
        TxtServer.Text = config.Server
        TxtDatabase.Text = config.Database
        CboAuth.Text = config.Authentication
    End Sub

    Private Async Sub BtnTest_Click(sender As Object, e As RoutedEventArgs)
        Dim testConfig = GetCurrentConfigFromUI()
        BtnTest.IsEnabled = False
        Dim success = Await _dbService.TestConnectionAsync(testConfig)
        BtnTest.IsEnabled = True

        If success Then
            MessageBox.Show("Connection successful!", "Test Connection", MessageBoxButton.OK, MessageBoxImage.Information)
        Else
            MessageBox.Show("Connection failed. Please check your settings.", "Test Connection", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        UpdatedConfig = GetCurrentConfigFromUI()
        DialogResult = True
        Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        DialogResult = False
        Close()
    End Sub

    Private Function GetCurrentConfigFromUI() As DatabaseConfig
        Return New DatabaseConfig() With {
            .Server = TxtServer.Text,
            .Database = TxtDatabase.Text,
            .Authentication = CboAuth.Text,
            .Username = "",
            .Password = ""
        }
    End Function
End Class
