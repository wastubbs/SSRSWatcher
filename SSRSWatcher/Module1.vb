Imports System
Imports System.IO
Imports System.Threading
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data
Imports System.Data.SqlClient

Module Module1
    Dim iUnProcessed As Integer
    Dim sConsolePath1 As String = _
        "C:\MSP_SSRS\"
    Dim sConsolepath2 As String = _
        "\SSRSConsole.exe"
    Dim sConnectionString As String = _
        "Server=ARETHA\TPI_SQL01;Database=MobileFrame;User Id=mspViewer;Password=Data4theroad!;"

    Sub Main()
        'GetCount("Server=ARETHA\TPI_SQL01; Database=MobileFrame; User Id=mspViewer; Password=Data4theroad!")
        '// The connection specified above wouldn't process. Had to create a 32-BIT ODBC connection, named below.
        '// **** This should be parameterized into a config file later on.
        While 1 = 1
            GetUnprocessedCount(sConnectionString)
            Console.Clear()
            Console.WriteLine("Current unprocessed jobs: " & iUnProcessed)
            'Console.ReadKey()
            '// **** This needs to be parameterized in a config file later on.
            If iUnProcessed = 0 Then
                Thread.Sleep(1000)
            Else
                GetRequestedJobs(sConnectionString)
            End If
        End While
    End Sub
    Public Sub GetRequestedJobs(ByVal sConnectionString As String)
        '// This is only used to retrieve unprocessed jobs.
        '// It should release the SQL connection as soon as it's done so 
        '// that it isn't kept open during processing.
        Dim queryString As String = _
            "SELECT * FROM SYS_JOBQUEUE WHERE JQPROCESSED = 0"
        Dim pConsole As New ProcessStartInfo
        Dim tdStarted, tdCompleted, sMessage As String

        Using objConn As New SqlConnection(sConnectionString)
            Dim cmd As New SqlCommand(queryString, objConn)
            cmd.Connection = objConn
            objConn.Open()

            Dim daJobs As New SqlDataAdapter(queryString, objConn)

            Dim dsJobs As New DataSet("Jobs")
            daJobs.FillSchema(dsJobs, SchemaType.Source, "SYS_JOBQUEUE")
            daJobs.Fill(dsJobs, "SYS_JOBQUEUE")

            Dim tblSYS_JOBQUEUE As DataTable
            tblSYS_JOBQUEUE = dsJobs.Tables("SYS_JOBQUEUE")

            Dim drCurrent As DataRow

            '// Begin processing the jobs...
            For Each drCurrent In tblSYS_JOBQUEUE.Rows
                If File.Exists(sConsolePath1 & drCurrent("JQUSER") & sConsolepath2) Then
                    tdStarted = Now()
                    Select Case drCurrent("JQJOB_TYPE")
                        Case "INTERNAL"
                            pConsole.FileName = sConsolePath1 & drCurrent("JQUSER") & sConsolepath2
                            pConsole.Arguments = _
                                "/url=http://172.22.1.13:8044/reportserver/ReportExecution2005.asmx /report=/WO_" & drCurrent("JQJOB_TYPE") & " /render=PDF /saveas=" & drCurrent("JQPATHNAME") & drCurrent("JQFILENAME") & " /p1=WONUM:" & drCurrent("JQKEY_1")
                            pConsole.UseShellExecute = True
                            pConsole.WindowStyle = ProcessWindowStyle.Normal


                            Dim proc As Process = Process.Start(pConsole)
                            proc.WaitForExit()
                            tdCompleted = Now()
                            If proc.ExitCode = 0 Then
                                sMessage = "SSRS: Job Complete"
                                UpdateJob(tdStarted, tdCompleted, True, sMessage, drCurrent("id"))
                            Else
                                sMessage = "Fail: SSRSConsole returned Exit Code (" & proc.ExitCode & ")"
                                UpdateJob(tdStarted, tdCompleted, True, sMessage, drCurrent("id"))
                            End If
                            Console.WriteLine(pConsole.Arguments)
                            'proc.Start(pConsole)
                            'Console.WriteLine("{0} {1}", drCurrent("deleted_by").ToString, drCurrent("JQJOB_ID").ToString)
                        Case Else
                            Console.WriteLine("This JQJOB_TYPE/Report Definition type has not yet been defined.")
                    End Select
                Else
                    Console.WriteLine(sConsolePath1 & drCurrent("JQUSER") & sConsolepath2 & " was not found!")
                End If
                Console.WriteLine(drCurrent("id"))
                'Console.ReadKey()
            Next
            'Console.ReadLine()


        End Using
        'Console.ReadKey()
    End Sub
    Private Sub GetUnprocessedCount(ByVal sConnectionString As String)
        '// This is only used to retrieve unprocessed jobs.
        '// It should release the SQL connection as soon as it's done so 
        '// that it isn't kept open during processing.
        Dim queryString As String = _
            "SELECT COUNT(*) FROM SYS_JOBQUEUE WHERE JQPROCESSED = 0"

        Using objConn As New SqlConnection(sConnectionString)
            Dim cmd As New SqlCommand(queryString, objConn)
            cmd.Connection = objConn
            objConn.Open()
            iUnProcessed = cmd.ExecuteScalar()
        End Using
    End Sub
    'Private Sub GetRequestedJobs(ByVal connectionString As String)
    '    Dim queryString As String = _
    '"SELECT * FROM SYS_JOBQUEUE WHERE JQ_PROCESSED = 0"
    '    Dim connection As New OdbcConnection(connectionString)
    '    Dim command As New OdbcCommand(queryString)
    'End Sub

    'Private Sub GetUnprocessedCount(ByVal connectionString As String)
    '    '// This is only used to retrieve unprocessed jobs
    '    Dim queryString As String = _
    '        "SELECT COUNT(*) FROM SYS_JOBQUEUE WHERE JQPROCESSED = 0"
    '    Dim command As New OdbcCommand(queryString)

    '    Using connection As New OdbcConnection(connectionString)
    '        command.Connection = connection
    '        connection.Open()
    '        iUnProcessed = command.ExecuteScalar()
    '    End Using
    'End Sub

    Private Sub UpdateJob(tdStarted As String, tdCompleted As String, p3 As Boolean, sMessage As String, id As Integer)
        sMessage = Left(sMessage, 200)
        sMessage = Replace(sMessage, "'", "''")

        Dim sQuery As String = "UPDATE SYS_JOBQUEUE " & _
                               "SET JQPROCESSED='1'," & _
                               "JQSTARTED='" & tdStarted & "'," & _
                               "JQMESSAGE='" & sMessage & "'," & _
                               "JQCOMPLETED='" & tdCompleted & "' " & _
            "WHERE [ID]='" & id & "'"

        Using objConn As New SqlConnection(sConnectionString)
            Dim cmd As New SqlCommand(sQuery, objConn)
            cmd.Connection = objConn
            objConn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub



End Module
