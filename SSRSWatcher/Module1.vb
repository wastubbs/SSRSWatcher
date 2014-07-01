Imports System
Imports System.IO
Imports System.Web.Services.Protocols
Imports System.Runtime.InteropServices
Imports System.Web.Services.Description
Imports SSRSConsole.SSRS_ReportExecution2005
Imports System.Data.OleDb

Module Module1

    Public NotInheritable Class OdbcConnection _
        Inherits DbConnection _
	    Implements ICloneable
    End Class


    Sub Main()
    End Sub

    Private Sub InsertRow(ByVal connectionString As String)

        Dim queryString As String = _
            "SELECT COUNT(*) FROM SYS_JOBQUEUE WHERE JQPROCESSED = 0"
        Dim command As New OdbcCommand(queryString)

        Using connection As New OdbcConnection(connectionString)
            command.Connection = connection
            connection.Open()
            command.ExecuteNonQuery()

            ' The connection is automatically closed at  
            ' the end of the Using block. 
        End Using
    End Sub

End Module
