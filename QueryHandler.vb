Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms


Public Class QueryHandler
    Private Shared _sqlConnection As SqlConnection
    Private Shared _sqlCommand As SqlCommand
    Private Shared _sqlDataAdapter As SqlDataAdapter
    Private Shared _dataTable As DataTable

    Public Function ValidateConnection(connectionString As String) As Boolean
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Return connection.State = ConnectionState.Open
            End Using
        Catch ex As Exception
            LogError(ex, displayError:=True)
            Return False
        End Try
    End Function

    Public Function ExecuteQuery(query As String, connectionString As String, Optional returnSingleRow As Boolean = True, Optional timeout As Integer = 3000) As Object
        Try
            If Not ValidateConnection(connectionString) Then
                LogMessage("Failed to establish connection. Please validate the server connection and try again.", displayMessage:=True)
                Return Nothing
            End If

            _sqlConnection = New SqlConnection(connectionString)
            _sqlCommand = New SqlCommand(query, _sqlConnection) With {
                .CommandType = CommandType.Text,
                .CommandTimeout = timeout
            }
            _sqlConnection.Open()
            _sqlCommand.ExecuteNonQuery()

            _sqlDataAdapter = New SqlDataAdapter(_sqlCommand)
            _dataTable = New DataTable()
            _sqlDataAdapter.Fill(_dataTable)

            If returnSingleRow Then
                If _dataTable.Rows.Count > 0 Then
                    Return _dataTable.Rows(0)
                End If
            Else
                Return _dataTable
            End If

            _sqlConnection.Close()
            _sqlConnection.Dispose()
            _sqlCommand.Dispose()
        Catch ex As Exception
            LogError(ex, displayError:=True)
            Return Nothing
        End Try

        Return Nothing
    End Function

    Public Sub ResetID(tableName As String, connectionString As String)
        ExecuteQuery($"DBCC CHECKIDENT('{tableName}')", connectionString, False)
    End Sub

    Public Function ExecuteQueryToDataTable(query As String, connectionString As String) As DataTable
        Try
            Dim conn As New SqlConnection(connectionString)
            conn.Open()
            Dim cmd As New SqlCommand(query, conn)

            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            conn.Close()
            Return dt
        Catch ex As Exception
            LogError(ex, displayError:=True)
            Return Nothing
        End Try
    End Function

    Public Function ExecuteQueryToDataRow(query As String, connectionString As String, Optional timeout As Integer = 0) As DataRow
        Try
            Dim conn As New SqlConnection(connectionString)
            conn.Open()
            Dim cmd As New SqlCommand(query, conn) With {
                .CommandTimeout = timeout
            }

            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            conn.Close()

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
            Return Nothing
        Catch ex As Exception
            LogError(ex, displayError:=True)
            Return Nothing
        End Try
    End Function

    Public Function ExecuteStoredProcedure(storedProcedureName As String, parameters As SqlParameter(), connectionString As String) As DataTable
        Try
            Using conn As New SqlConnection(connectionString)
                Using cmd As New SqlCommand(storedProcedureName, conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If

                    Dim dt As New DataTable()
                    conn.Open()
                    dt.Load(cmd.ExecuteReader())
                    conn.Close()
                    Return dt
                End Using
            End Using
        Catch ex As Exception
            LogError(ex, displayError:=True)
            Return Nothing
        End Try
    End Function

    Public Function ExecuteBatchQuery(query As String, parameters As SqlParameter(), connectionString As String) As Integer
        Try
            Using conn As New SqlConnection(connectionString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.CommandType = CommandType.Text
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If

                    conn.Open()
                    Dim affectedRows As Integer = cmd.ExecuteNonQuery()
                    conn.Close()
                    Return affectedRows
                End Using
            End Using
        Catch ex As Exception
            LogError(ex, displayError:=True)
            Return -1
        End Try
    End Function

    Private Sub LogError(ex As Exception, Optional displayError As Boolean = True, Optional title As String = "Error", Optional moduleName As String = "QueryHandler")
        ' Implement error logging logic here (e.g., write to a file, database, etc.)
        Dim logFilePath As String = "error_log.txt"
        Dim errorMessage As String = $"{DateTime.Now}: {moduleName} - {title}: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}"

        File.AppendAllText(logFilePath, errorMessage)

        If displayError Then
            LogMessage($"{title}: {ex.Message}", displayMessage:=True)
        End If
    End Sub

    Private Sub LogMessage(message As String, Optional displayMessage As Boolean = False)

        If displayMessage Then
            MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' Implement message logging logic here (e.g., write to a file, database, etc.)
            Dim logFilePath As String = "log.txt"
            Dim logMessage As String = $"{DateTime.Now}: {message}{Environment.NewLine}"

            File.AppendAllText(logFilePath, logMessage)

        End If
    End Sub
End Class

