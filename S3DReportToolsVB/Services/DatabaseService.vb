Imports System.Data
Imports Microsoft.Data.SqlClient
Imports S3DReportToolsVB.Models

Namespace Services
    Public Class DatabaseService
        Public Async Function ExecuteQueryAsync(config As DatabaseConfig, query As String) As Task(Of DataTable)
            Dim dt As New DataTable()
            Using conn As New SqlConnection(config.GetConnectionString())
                Await conn.OpenAsync()
                Using cmd As New SqlCommand(query, conn)
                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using
            Return dt
        End Function

        Public Async Function TestConnectionAsync(config As DatabaseConfig) As Task(Of Boolean)
            Try
                Using conn As New SqlConnection(config.GetConnectionString())
                    Await conn.OpenAsync()
                    Return True
                End Using
            Catch
                Return False
            End Try
        End Function
    End Class
End Namespace
