Imports System.Data.SqlClient
Imports System.Configuration

Public Module DBConnection


    Private Const OpticTestConn As String = "Data Source=;User ID=;Password=;Initial Catalog=OptimaxDeveloperTest;Application Name=OptimaxTest.Service"

    Private Sub LogError(ByVal ex As Exception)
        LogError("", ex)
    End Sub

    Private Sub LogError(ByVal ex As Exception, ByVal Details As String)
        LogError("", ex, Details)
    End Sub

    Private Sub LogError(ByVal ex As Exception, ByVal argSqlCommand As SqlClient.SqlCommand)
        LogError("", ex, argSqlCommand)
    End Sub

    Public Sub LogError(ByVal ClientDetails As String, ByVal ex As Exception)

        Dim aLog As EventLog
        Dim strClient As String

        aLog = New EventLog
        aLog.Source = "Optic"

        If ClientDetails = "" Then
            strClient = ""
        Else
            strClient = "Client Details: " & ClientDetails & ControlChars.CrLf
        End If

        aLog.WriteEntry(strClient & ex.Message & ControlChars.CrLf & "Stack Trace" & ex.StackTrace, EventLogEntryType.Error)

        Throw ex 'pass exception back up 

    End Sub

    Public Sub LogError(ByVal ClientDetails As String, ByVal ex As Exception, ByVal argSQLCommand As SqlClient.SqlCommand)
        'added to allow passing in of the command object
        Dim aLog As EventLog
        Dim strClient As String
        Dim strError As String

        strError = ""

        Try
            If argSQLCommand Is Nothing Then
                strError = "SqlClient.SqlCommand object is nothing." & ControlChars.NewLine

            ElseIf argSQLCommand.Parameters Is Nothing Then
                strError = "SqlClient.SqlCommand.Parameters object has not been created." & ControlChars.NewLine

            ElseIf argSQLCommand.Parameters.Count < 1 Then
                strError = "SqlClient.SqlCommand.Parameters object has no parameters." & ControlChars.NewLine

            Else
                strError += argSQLCommand.CommandType.ToString().PadRight(44) & ControlChars.Tab & " =" &
                   "[" & argSQLCommand.CommandText.ToString() & "]"

                For Each pError As SqlClient.SqlParameter In argSQLCommand.Parameters
                    strError += pError.ParameterName.ToString.PadRight(30) & ControlChars.Tab &
                       ("[" & pError.SqlDbType.ToString() & "]").PadRight(14) & ControlChars.Tab & " =" &
                       ControlChars.Tab & pError.Value.ToString() & ControlChars.NewLine
                Next
            End If

        Catch x As Exception
            strError = ""
        End Try

        aLog = New EventLog
        aLog.Source = "Optic"

        If ClientDetails = "" Then
            strClient = ""
        Else
            strClient = "Client Details: " & ClientDetails & ControlChars.CrLf
        End If

        aLog.WriteEntry(ex.Message & ControlChars.CrLf & "Stack Trace" & ex.StackTrace & ControlChars.NewLine & strError, EventLogEntryType.Error)

        Throw ex

    End Sub

    Public Sub LogError(ByVal ClientDetails As String, ByVal ex As Exception, ByVal strDetail As String)
        'added to allow extra detail in debug - line no's/objects etc
        Dim aLog As EventLog
        Dim strClient As String

        aLog = New EventLog
        aLog.Source = "Optic"

        If ClientDetails = "" Then
            strClient = ""
        Else
            strClient = "Client Details: " & ClientDetails & ControlChars.CrLf
        End If

        aLog.WriteEntry(ex.Message & ControlChars.CrLf & "Stack Trace" & ex.StackTrace & ControlChars.NewLine & strDetail, EventLogEntryType.Error)

        Throw ex

    End Sub



    Public Enum SQLConnectionStrings
        OPTIC = 1
        OptimaxReports = 2
    End Enum



    Public Function MakeConnection() As SqlClient.SqlConnection
        Dim sqcLocal As SqlClient.SqlConnection

        sqcLocal = Nothing

        Try
            sqcLocal = New SqlClient.SqlConnection
            sqcLocal.ConnectionString = OpticTestConn
            sqcLocal.Open()


        Catch sqlex As SqlException
            LogError(sqlex)

        Catch ex As Exception
            LogError(ex)

        End Try

        Return sqcLocal

    End Function

    Public Function MakeConnection(ByVal pConnectTo As SQLConnectionStrings) As SqlClient.SqlConnection
        Dim sqcLocal As SqlClient.SqlConnection

        sqcLocal = Nothing

        Try
            ' Ok this uses pooled connections via the connection string, 
            ' ie if the same then ADO uses existing connection if 
            ' idle or pool has reached max connections. Otherwise a new connection will be generated 
            ' but only on needed.

            sqcLocal = New SqlClient.SqlConnection

            sqcLocal.ConnectionString = OpticTestConn

            sqcLocal.Open()

        Catch sqlex As SqlException
            LogError(sqlex)

        Catch ex As Exception
            LogError(ex)

        End Try

        Return sqcLocal

    End Function

    Private Sub CloseConnection(ByRef argConnection As SqlClient.SqlConnection)

        Try
            ' closes the pooled connection
            argConnection.Close()

        Catch sqlex As SqlException
            LogError(sqlex)

        Catch ex As Exception
            LogError(ex)

        End Try

    End Sub

    Public Sub CloseAndDisposeConnection(ByRef argConnection As SqlClient.SqlConnection)

        Try
            ' closes the pooled connection
            argConnection.Close()

            'disposes the pooled collection
            argConnection.Dispose()

            'forces disposal
            ' GC.Collect()
            ' Updated by Richard Russell - Long Term Garage Collection Needed Using deep clean command
            GC.Collect(2)

        Catch sqlex As SqlException
            LogError(sqlex)

        Catch ex As Exception
            LogError(ex)

        End Try

    End Sub

    Private Function GetSQLServerTimeStamp() As Date
        Dim sqcLocal As SqlClient.SqlConnection
        Dim sql As String
        Dim scLocal As SqlCommand
        Dim sdaLocal As SqlDataAdapter
        Dim dsLocal As DataSet
        Dim StrDetail As String
        Dim timeStamp As Date

        timeStamp = Nothing
        StrDetail = ""
        sqcLocal = Nothing

        Try
            ' Make sure there is a pooled connection to the database
            StrDetail = "MakeConnection()" & ControlChars.NewLine & "line 1"
            sqcLocal = MakeConnection()

            ' This is the T-SQL which will return the system datetime
            sql =
             "SELECT" & ControlChars.NewLine & ControlChars.Tab &
              "GetDate() AS [TimeStamp]"

            ' Create a new SqlCommand using the T-SQL statement and the
            ' pooled connection to the database
            StrDetail = "scModule" & ControlChars.NewLine & "line 2"
            scLocal = New SqlCommand(sql, sqcLocal)

            ' Richard Russell 08/05/2005
            ' Command to increase timeout for all SQL statements to allow for longer running commands 
            ' 600 sec= 5 minutes for all those without a calculator.
            ' Really this should passed as parameter from  the calling routine, 
            ' but this is meant as a quick fix.

            ' Steve Davis 09/05/2005
            ' I don't know where you buy your calculators from Richard!
            ' 300 sec = 5 minutes
            ' I have changed this so it is in line with other timeouts in this module
            scLocal.CommandTimeout = 300 ' 5 Minutes 

            ' Create a new SqlDataAdapter using the SqlCommand
            StrDetail = "scLocal" & ControlChars.NewLine & "line 3"
            sdaLocal = New SqlDataAdapter(scLocal)

            ' Create a new DataSet ready for population
            dsLocal = New DataSet

            ' Populate the DataSet
            StrDetail = "sdaLocal.Fill(dsLocal)" & ControlChars.NewLine & "line 4"
            sdaLocal.Fill(dsLocal)

            ' Return the datetime value from the dataset
            StrDetail = "dsLocal.Tables(0).Rows(0).Item(TimeStamp)" & ControlChars.NewLine & "line 5"

            timeStamp = Convert.ToDateTime(dsLocal.Tables(0).Rows(0).Item("TimeStamp"))

        Catch sqlex As SqlException
            LogError(sqlex, StrDetail)

        Catch ex As Exception
            LogError(ex, StrDetail)

        Finally
            CloseConnection(sqcLocal)

        End Try

        Return timeStamp

    End Function

    'Public Function GetDatabaseIPAddress() As String

    '    Dim strIPAddress As String
    '    Dim hostInfo As System.Net.IPHostEntry

    '    strIPAddress = ""

    '    Try
    '        'All above lines handled by line below
    '        hostInfo = System.Net.Dns.GetHostEntry("OpticDBServer")

    '        ' Get Server IP Address
    '        strIPAddress = hostInfo.AddressList(0).ToString()

    '    Catch ex As Exception
    '        LogError(ex)

    '    End Try

    '    Return strIPAddress

    'End Function

    Public Function GetDataSet(ByVal SQLCmd As String, ByVal pConnection As SQLConnectionStrings) As DataSet
        Dim sqcLocal As SqlClient.SqlConnection
        Dim dLocal As Date
        Dim scLocal As SqlCommand
        Dim sdaLocal As SqlDataAdapter
        Dim dsLocal As DataSet
        Dim dtLocal As DataTable
        Dim dcLocal As DataColumn
        Dim drLocal As DataRow
        Dim strdetail As String
        Dim blnShowSQL As Boolean

        strdetail = ""
        sqcLocal = Nothing
        dsLocal = Nothing

        Try
            blnShowSQL = Convert.ToBoolean(ConfigurationManager.AppSettings("ReturnSQL"))

            ' Get the SQL Server datetime
            If blnShowSQL Then
                strdetail = "GetSQLServerTimeStamp()" & ControlChars.NewLine & "line 1"
                dLocal = GetSQLServerTimeStamp()
            End If

            ' Ensure there is a pooled connection to the database
            strdetail = "MakeConnection()" & ControlChars.NewLine & "line 2"
            sqcLocal = MakeConnection(pConnection)

            ' Create a new SQLCommand using the supplied T-SQL statement and the
            ' pooled connection to the database
            strdetail = "scModule" & ControlChars.NewLine & "line 3"
            scLocal = New SqlCommand(SQLCmd, sqcLocal)

            ' Set the timeout to 5 minutes
            scLocal.CommandTimeout = 300

            ' Create a new SQLDataAdapter using the SQLCommand
            strdetail = "scLocal" & ControlChars.NewLine & "line 4"
            sdaLocal = New SqlDataAdapter(scLocal)

            ' Create a new DataSet ready for population, which will be returned
            ' by this function
            dsLocal = New DataSet

            ' Populate the DataSet
            If IsNothing(dsLocal) Then
                strdetail = "dsLocal is nothing" & ControlChars.NewLine & "line 5"
            ElseIf IsNothing(sdaLocal) Then
                strdetail = "sdaLocal is nothing" & ControlChars.NewLine & "line 5"
            ElseIf IsNothing(sqcLocal) Then
                strdetail = "scModule is nothing" & ControlChars.NewLine & "line 5"
            ElseIf IsNothing(scLocal) Then
                strdetail = "scLocal is nothing" & ControlChars.NewLine & "line 5"
            Else
                strdetail = "dsLocal/sdalocal/scModule/sclocal neither are nothing" & ControlChars.NewLine & "line 5"
            End If
            sdaLocal.Fill(dsLocal)

            If blnShowSQL = True Then

                ' Attach an additional table to the DataSet for system management information
                ' so it's available for updates
                strdetail = "dsLocal.tables.add" & ControlChars.NewLine & "line 6"
                dtLocal = dsLocal.Tables.Add("OPTIC System Management Information")

                ' Add a column for the SQL
                strdetail = "dtLocal.Columns.Add(SQL)" & ControlChars.NewLine & "line 7"
                dcLocal = dtLocal.Columns.Add("SQL")

                ' Add a column for the TimeStamp
                strdetail = "dtLocal.Columns.Add(TimeStamp)" & ControlChars.NewLine & "line 8"
                dcLocal = dtLocal.Columns.Add("TimeStamp")

                ' Create a new row from the system management information table
                strdetail = "dtLocal.NewRow" & ControlChars.NewLine & "line 9"
                drLocal = dtLocal.NewRow()

                ' Populate the row with system management information
                drLocal("SQL") = SQLCmd

                strdetail = "dLocal" & ControlChars.NewLine & "line 10"
                drLocal("Timestamp") = dLocal

                ' Add the system management information row to the system management
                ' information table
                strdetail = "drLocal" & ControlChars.NewLine & "line 11"
                dtLocal.Rows.Add(drLocal)

            End If

        Catch sqlex As SqlException
            LogError(sqlex, strdetail)

        Catch ex As Exception
            LogError(ex, strdetail)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

        Return dsLocal

    End Function

    Public Function GetDataSet(ByRef argSQLCommand As SqlCommand) As DataSet

        Return GetDataSet(argSQLCommand,
                Convert.ToBoolean(ConfigurationManager.AppSettings("ReturnSQL")),
                SQLConnectionStrings.OPTIC)


    End Function

    Public Function GetDataSet(ByVal pSQLCommand As SqlCommand, ByVal pConnection As SQLConnectionStrings) As DataSet

        Return GetDataSet(pSQLCommand,
                    Convert.ToBoolean(ConfigurationManager.AppSettings("ReturnSQL")),
                    pConnection)

    End Function

    Public Function GetDataSet(ByVal pSQL As String) As DataSet

        Return GetDataSet(pSQL, SQLConnectionStrings.OPTIC)

    End Function

    Public Function GetDataSetFromSPWithNoParams(ByVal pStoredProcName As String) As DataSet
        Dim sqlCmd As SqlClient.SqlCommand
        Dim myData As DataSet

        myData = Nothing

        sqlCmd = New SqlClient.SqlCommand
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.CommandText = pStoredProcName

        Try
            myData = GetDataSet(sqlCmd)

        Catch ex As Exception
            LogError(ex)

        End Try

        Return myData

    End Function

    Public Function GetDataSet(ByRef argSQLCommand As SqlCommand,
                    ByVal argAddSystemManagementInformation As Boolean,
                    ByVal pConnection As SQLConnectionStrings) As DataSet

        Dim sqcLocal As SqlClient.SqlConnection
        Dim dLocal As Date
        Dim sdaLocal As SqlDataAdapter
        Dim dsLocal As DataSet
        Dim dtLocal As DataTable
        Dim dcLocal As DataColumn
        Dim drLocal As DataRow

        sqcLocal = Nothing
        dsLocal = Nothing

        Try
            ' Get the SQL Server datetime
            If argAddSystemManagementInformation = True Then
                dLocal = GetSQLServerTimeStamp()
            End If

            ' Ensure there is a pooled connection to the database
            sqcLocal = MakeConnection(pConnection)

            ' Ensure the supplied SQLCommand uses the pooled database connection
            argSQLCommand.Connection = sqcLocal

            ' Set the timeout to 5 minutes
            argSQLCommand.CommandTimeout = 300

            ' Create a new SQLDataAdapter using the supplied SQLCommand
            sdaLocal = New SqlDataAdapter(argSQLCommand)

            ' Create a new DataSet ready for population, which will be returned
            ' by this function
            dsLocal = New DataSet

            ' Populate the DataSet
            sdaLocal.Fill(dsLocal)

            If argAddSystemManagementInformation = True Then
                ' Attach an additional table to the DataSet for system management information
                ' so it's available for updates
                dtLocal = dsLocal.Tables.Add("OPTIC System Management Information")

                dcLocal = dtLocal.Columns.Add("SQL")
                dcLocal = dtLocal.Columns.Add("TimeStamp")

                ' Create a new row from the system management information table
                drLocal = dtLocal.NewRow()

                ' Populate the row with system management information
                drLocal("SQL") = argSQLCommand.CommandText
                drLocal("TimeStamp") = dLocal

                ' Add the system mangement information row to the syatem mangement 
                ' information table
                dtLocal.Rows.Add(drLocal)
            End If

        Catch sqlex As SqlException
            LogError(sqlex, argSQLCommand)

        Catch ex As Exception
            LogError(ex, argSQLCommand)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

        Return dsLocal

    End Function

    Public Function GetReader(ByVal argSQLCommand As SqlCommand) As SqlDataReader
        argSQLCommand.Connection = MakeConnection()
        Dim reader As SqlDataReader = argSQLCommand.ExecuteReader()
        Return reader
    End Function

    Public Function GetReaderValue(ByVal myReader As SqlDataReader, ByVal Column As String) As String
        If IsDBNull(myReader.GetValue(myReader.GetOrdinal(Column))) Then
            Return ""
        Else
            Return myReader.GetString(myReader.GetOrdinal(Column))
        End If
    End Function

    Public Sub ExecuteSql(ByVal pSqlCmd As SqlClient.SqlCommand, ByVal pConnection As SQLConnectionStrings)
        Dim sqcLocal As SqlClient.SqlConnection

        sqcLocal = Nothing

        Try
            sqcLocal = MakeConnection(pConnection)

            pSqlCmd.Connection = sqcLocal
            pSqlCmd.ExecuteNonQuery()

        Catch sqlex As SqlException
            LogError(sqlex, pSqlCmd)

        Catch ex As Exception
            LogError(ex)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

    End Sub

    Public Sub ExecuteSQL(ByRef SqlCmd As SqlClient.SqlCommand)
        Dim sqcLocal As SqlClient.SqlConnection

        sqcLocal = Nothing

        Try
            sqcLocal = MakeConnection()

            SqlCmd.Connection = sqcLocal
            SqlCmd.ExecuteNonQuery()

        Catch sqlex As SqlException
            LogError(sqlex, SqlCmd)

        Catch ex As Exception
            LogError(ex)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

    End Sub

    Public Sub ExecuteSql(ByVal pSqlCmd As String, ByVal pConnection As SQLConnectionStrings)
        Dim sqcLocal As SqlClient.SqlConnection
        Dim scLocal As SqlCommand

        sqcLocal = Nothing

        Try
            sqcLocal = MakeConnection(pConnection)

            scLocal = New SqlCommand(pSqlCmd, sqcLocal)
            scLocal.ExecuteNonQuery()

        Catch sqlex As SqlException
            LogError(sqlex, pSqlCmd)

        Catch ex As Exception
            LogError(ex)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

    End Sub

    Public Sub ExecuteSql(ByVal pSqlCmd As String)
        Dim sqcLocal As SqlClient.SqlConnection
        Dim scLocal As SqlCommand

        sqcLocal = Nothing

        Try
            sqcLocal = MakeConnection()

            scLocal = New SqlCommand(pSqlCmd, sqcLocal)
            scLocal.ExecuteNonQuery()

        Catch sqlex As SqlException
            LogError(sqlex, pSqlCmd)

        Catch ex As Exception
            LogError(ex)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

    End Sub

    Public Sub ExecuteLongRunningSQL(ByRef SqlCmd As SqlClient.SqlCommand)
        Dim sqcLocal As SqlClient.SqlConnection

        sqcLocal = Nothing

        Try
            sqcLocal = MakeConnection()

            SqlCmd.Connection = sqcLocal
            SqlCmd.CommandTimeout = 300 '5 minutes
            SqlCmd.ExecuteNonQuery()

        Catch sqlex As SqlException
            LogError(sqlex, SqlCmd)

        Catch ex As Exception
            LogError(ex)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

    End Sub

    Public Sub ExecuteLongRunningSQL(ByRef SqlCmd As SqlClient.SqlCommand, ByVal pConnection As SQLConnectionStrings)
        Dim sqcLocal As SqlClient.SqlConnection

        sqcLocal = Nothing

        Try
            sqcLocal = MakeConnection(pConnection)

            SqlCmd.Connection = sqcLocal
            SqlCmd.CommandTimeout = 300 '5 minutes
            SqlCmd.ExecuteNonQuery()

        Catch sqlex As SqlException
            LogError(sqlex, SqlCmd)

        Catch ex As Exception
            LogError(ex)

        Finally
            CloseAndDisposeConnection(sqcLocal)

        End Try

    End Sub

End Module
