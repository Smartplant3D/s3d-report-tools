Namespace Models
    Public Class DatabaseConfig
        Public Property Server As String = "SPF"
        Public Property Database As String = "GDP34_RDB"
        Public Property Authentication As String = "Windows Authentication"
        Public Property Username As String = ""
        Public Property Password As String = ""

        Public Function GetConnectionString() As String
            ' Always use Windows Authentication as requested
            Return $"Server={Server};Database={Database};Integrated Security=True;TrustServerCertificate=True;"
        End Function
    End Class
End Namespace
