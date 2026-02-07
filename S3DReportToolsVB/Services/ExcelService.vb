Imports System.Data
Imports ClosedXML.Excel
Imports System.IO

Namespace Services
    Public Class ExcelService
        Public Sub ExportToExcel(dt As DataTable, filePath As String)
            Using workbook As New XLWorkbook()
                Dim worksheet = workbook.Worksheets.Add(dt, "Results")
                ' Auto-fit columns to align data correctly and avoid text cutoff
                worksheet.Columns().AdjustToContents()
                workbook.SaveAs(filePath)
            End Using
        End Sub
    End Class
End Namespace
