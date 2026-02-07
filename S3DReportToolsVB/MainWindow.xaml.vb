Imports System.Data
Imports Microsoft.Win32
Imports S3DReportToolsVB.Models
Imports S3DReportToolsVB.Services

Class MainWindow
    Private Property Config As New DatabaseConfig()
    Private ReadOnly _dbService As New DatabaseService()
    Private ReadOnly _excelService As New ExcelService()

    Public Sub New()
        InitializeComponent()
        
        ' Load persisted settings or use defaults
        Config.Server = If(String.IsNullOrEmpty(My.Settings.Server), "SPF", My.Settings.Server)
        Config.Database = If(String.IsNullOrEmpty(My.Settings.Database), "GDP34_RDB", My.Settings.Database)
        
        Dim defaultSql = "select m.ItemName AreaName,j.ItemName UnitName,g.ItemName PipingSystem,e.ItemName PipeLineName,n.ItemName PipeRunName,b.ItemName ComponentsName,j1.ShortMaterialDescription,a.DateLastModified,ul.UserLogin UserLastModified" & vbCrLf &
                        "from JDObject a" & vbCrLf &
                        "join XPartOccToMaterialControlData x1 on x1.OidOrigin=a.Oid" & vbCrLf &
                        "join JGenericMaterialControlData j1 on j1.Oid = x1.OidDestination" & vbCrLf &
                        "join JUserLogin ul on ul.oid=UIDLastModifier" & vbCrLf &
                        "Join JNamedItem b on a.oid=b.oid" & vbCrLf &
                        "join XOwnsParts c on c.OidDestination=a.oid " & vbCrLf &
                        "join JNamedItem n on n.Oid=c.OidOrigin" & vbCrLf &
                        "Join XSystemHierarchy d on d.OidDestination=n.Oid" & vbCrLf &
                        "Join JNamedItem e on d.OidOrigin=e.Oid" & vbCrLf &
                        "Join XSystemHierarchy f on f.OidDestination=e.Oid" & vbCrLf &
                        "Join JNamedItem g on g.Oid=f.OidOrigin" & vbCrLf &
                        "Join XSystemHierarchy h on h.OidDestination=g.Oid" & vbCrLf &
                        "Join JNamedItem j on j.Oid=h.OidOrigin" & vbCrLf &
                        "Join XSystemHierarchy k on k.OidDestination=j.Oid" & vbCrLf &
                        "Join JNamedItem m on m.Oid=k.OidOrigin"

        TxtQuery.Text = If(String.IsNullOrEmpty(My.Settings.LastQuery), defaultSql, My.Settings.LastQuery)
    End Sub

    Private Async Sub BtnExecute_Click(sender As Object, e As RoutedEventArgs)
        Try
            BtnExecute.IsEnabled = False
            
            ' Save current query state
            My.Settings.LastQuery = TxtQuery.Text
            My.Settings.Save()

            Dim query = TxtQuery.Text
            Dim results = Await _dbService.ExecuteQueryAsync(Config, query)
            GridResults.ItemsSource = results.DefaultView
        Catch ex As Exception
            MessageBox.Show($"Error executing query: {ex.Message}", "Query Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            BtnExecute.IsEnabled = True
        End Try
    End Sub

    Private Sub BtnExport_Click(sender As Object, e As RoutedEventArgs)
        Try
            If GridResults.ItemsSource Is Nothing Then
                MessageBox.Show("Please execute a query first.", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            saveFileDialog.DefaultExt = "xlsx"
            saveFileDialog.FileName = $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"

            If saveFileDialog.ShowDialog() = True Then
                Dim dt = DirectCast(GridResults.ItemsSource, DataView).Table
                _excelService.ExportToExcel(dt, saveFileDialog.FileName)
                MessageBox.Show("Export successful!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error exporting: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnSettings_Click(sender As Object, e As RoutedEventArgs)
        Dim settingsWin As New ConnectionWindow(Config)
        If settingsWin.ShowDialog() = True Then
            Config = settingsWin.UpdatedConfig
            ' Persist the new settings
            My.Settings.Server = Config.Server
            My.Settings.Database = Config.Database
            My.Settings.Save()
        End If
    End Sub

    Private Sub BtnExit_Click(sender As Object, e As RoutedEventArgs)
        Application.Current.Shutdown()
    End Sub
End Class
