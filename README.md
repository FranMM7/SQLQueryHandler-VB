Certainly! Below is the `README` for the VB.NET version of the `SqlQueryHandler` project.

# SqlQueryHandler

SqlQueryHandler is a robust and efficient VB.NET library for handling SQL queries and stored procedures. This library provides functionalities for executing SQL queries, validating connections, executing stored procedures, and logging errors or messages.

## Features

- Validate SQL Server connections.
- Execute SQL queries and return results as `DataTable` or `DataRow`.
- Execute SQL stored procedures with parameters.
- Execute batch queries with parameters.
- Log errors and messages with an option to display them to the user.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/frankmejia7/QueryHandler.git
    ```

2. Open the project in Visual Studio.

3. Build the project to generate the `SqlQueryHandler.dll`.

4. Reference the `SqlQueryHandler.dll` in your .NET project.

## Usage

### 1. Validate Connection

```vb
Imports SqlQueryHandler

Module Program
    Sub Main()
        Dim queryHandler As New QueryHandler()
        Dim isConnected As Boolean = queryHandler.ValidateConnection("YourConnectionStringHere")

        If isConnected Then
            Console.WriteLine("Connection successful.")
        Else
            Console.WriteLine("Failed to connect.")
        End If
    End Sub
End Module
```

### 2. Execute SQL Query

```vb
Imports SqlQueryHandler

Module Program
    Sub Main()
        Dim queryHandler As New QueryHandler()
        Dim query As String = "SELECT * FROM YourTable"
        Dim result As Object = queryHandler.ExecuteQuery(query, "YourConnectionStringHere")

        If TypeOf result Is DataTable Then
            Dim dt As DataTable = DirectCast(result, DataTable)
            Console.WriteLine($"Rows returned: {dt.Rows.Count}")
        ElseIf TypeOf result Is DataRow Then
            Dim dr As DataRow = DirectCast(result, DataRow)
            Console.WriteLine($"First row data: {dr(0)}")
        End If
    End Sub
End Module
```

### 3. Execute Stored Procedure

```vb
Imports System.Data
Imports System.Data.SqlClient
Imports SqlQueryHandler

Module Program
    Sub Main()
        Dim queryHandler As New QueryHandler()
        Dim storedProcedure As String = "YourStoredProcedureName"
        Dim parameters() As SqlParameter = {
            New SqlParameter("@ParameterName", SqlDbType.VarChar) With {
                .Value = "ParameterValue"
            }
        }

        Dim result As DataTable = queryHandler.ExecuteStoredProcedure(storedProcedure, parameters, "YourConnectionStringHere")
        Console.WriteLine($"Stored Procedure returned: {result.Rows.Count} rows")
    End Sub
End Module
```

### 4. Execute Batch Query

```vb
Imports System.Data
Imports System.Data.SqlClient
Imports SqlQueryHandler

Module Program
    Sub Main()
        Dim queryHandler As New QueryHandler()
        Dim query As String = "UPDATE YourTable SET Column1 = @Param1 WHERE Column2 = @Param2"
        Dim parameters() As SqlParameter = {
            New SqlParameter("@Param1", SqlDbType.Int) With {
                .Value = 1
            },
            New SqlParameter("@Param2", SqlDbType.VarChar) With {
                .Value = "Value"
            }
        }

        Dim rowsAffected As Integer = queryHandler.ExecuteBatchQuery(query, parameters, "YourConnectionStringHere")
        Console.WriteLine($"Batch query affected: {rowsAffected} rows")
    End Sub
End Module
```

### 5. Log Messages and Errors

```vb
Imports SqlQueryHandler

Module Program
    Sub Main()
        Dim queryHandler As New QueryHandler()

        ' Log a simple message
        queryHandler.LogMessage("This is a log message.", displayMessage:=True)

        ' Log an error
        Try
            ' Some code that throws an exception
        Catch ex As Exception
            queryHandler.LogError(ex, displayError:=True)
        End Try
    End Sub
End Module
```

## Contributing

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add some feature'`).
5. Push to the branch (`git push origin feature-branch`).
6. Create a new Pull Request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

This `README` provides a comprehensive guide for using the SqlQueryHandler library in VB.NET. If you have any issues or questions, feel free to open an issue on the GitHub repository or contact the maintainer.
