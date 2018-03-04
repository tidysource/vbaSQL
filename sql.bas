Attribute VB_Name = "VBA_require_SQL"
Option Explicit

Function setupSQL( _
    dbPath As String, _
    DDLfolderPath As String) _
    As Boolean
    'Connect to database
    '-------------------
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open ( _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbPath & ";" & _
        "User Id=admin;Password=" _
        )

    'Perform the SQL query
    '---------------------
    conn.BeginTrans

    'Read folder
    Dim files As Variant
    files = getFiles(DDLfolderPath)

    'For each file
    Dim i As Integer
    Dim SQL As String
    Dim regex As Object
    Dim filePath As String
    Set regex = CreateObject("VBScript.RegExp")
    For i = LBound(files) To UBound(files)
        filePath = CStr(files(i))
        If pathExt(filePath) = ".sql" Then
            'Read file
            SQL = readFile(filePath)

            'Remove comments
            With regex
                .Pattern = "\/\*[\s\S]*?\*\/"
                .Global = True
            End With
            SQL = regex.Replace(SQL, "")

            With regex
                .Pattern = "--.*"
                .Global = True
            End With
            SQL = regex.Replace(SQL, "")

            'Constraints of access
            With regex
                .Pattern = "BIGINT"
                .Global = True
            End With
            SQL = regex.Replace(SQL, "INTEGER")

            With regex
                .Pattern = "INTEGER UNSIGNED"
                .Global = True
            End With
            SQL = regex.Replace(SQL, "INTEGER")

            With regex
                .Pattern = "NUMERIC"
                .Global = True
            End With
            SQL = regex.Replace(SQL, "DOUBLE")

            'Debug.Print filePath
            Debug.Print SQL
            conn.Execute SQL
        End If
    Next i

    'Commit transaction
    conn.CommitTrans

    'Close database connection
    '-------------------------
    conn.Close

    setupSQL = True
Exit Function
transError:
    conn.RollBack
    conn.Close
    setupSQL = False
    Debug.Print err.Description
End Function

'Perform a single SQL statement
'------------------------------
'For single SQL statement we use this,
'otherwise SQLtransact will work too,
'although a bit slower
Function SQLexec( _
    ByVal dbPath As String, _
    ByVal sqlStatement As String _
    ) As Boolean

    'Connect to database
    '-------------------
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open ( _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbPath & ";" & _
        "User Id=admin;Password=" _
        )

    'Perform the SQL queries
    '-----------------------
    conn.Execute sqlStatement

    'Close database connection
    '-------------------------
    conn.Close

    SQLexec = True
End Function

'Perform one or more SQL statements
'----------------------------------
Function SQLtransact( _
    ByVal dbPath As String, _
    ParamArray statements() As Variant _
    ) As Boolean

    SQLtransact = SQLtransaction(dbPath, statements)
End Function

'Perform one or more SQL statements
'----------------------------------
Function SQLtransaction( _
    ByVal dbPath As String, _
    ByVal statements As Variant _
    ) As Boolean

    'Connect to database
    '-------------------
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open ( _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbPath & ";" & _
        "User Id=admin;Password=" _
        )

    'Perform the SQL queries
    '-----------------------
    'Begin transaction
    conn.BeginTrans

    'For each SQL statement
    Dim i As Long
    Dim SQL As String
    For i = LBound(statements) To UBound(statements)
        SQL = CStr(statements(i))
        conn.Execute SQL
    Next i

    'Commit transaction
    conn.CommitTrans

    'Close database connection
    '-------------------------
    conn.Close

    SQLtransaction = True
Exit Function
transError:
    conn.RollBack
    conn.Close
    SQLtransaction = False
    Debug.Print err.Description
End Function

'Paste the result of an SELECT query into an Excel Sheet
'-------------------------------------------------------
Function SQLselect( _
    ByVal dbPath As String, _
    ByVal sqlStatement As String, _
    ByVal topLeftRow As Long, _
    ByVal topLeftColumn As Integer, _
    Optional ByVal sheetName As String = "", _
    Optional ByVal wbName As String = "", _
    Optional ByVal userPrompt As Boolean = True _
    ) As Boolean
    'Set default values
    If sheetName = "" Then
        sheetName = Application.ActiveWorkbook.ActiveSheet.Name
        wbName = Application.ActiveWorkbook.Name
    ElseIf wbName = "" Then
        wbName = Application.ActiveWorkbook.Name
    End If

    'Connect to database
    '-------------------
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open ( _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbPath & ";" & _
        "User Id=admin;Password=" _
        )

    'Perform the SQL queries
    '-----------------------
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    Set rs = conn.Execute(sqlStatement)

    'If SELECT returns any rows
    If rs.EOF <> True Then
        'Paste rows to sheet
        Application _
            .Workbooks(wbName) _
            .Sheets(sheetName) _
            .Cells(topLeftRow, topLeftColumn) _
                .CopyFromRecordset rs
    ElseIf userPrompt = True Then
        alert "No records found"
'        Debug.Print "-------------------------------" & _
'                    vbNewLine & _
'                    "No records match the SQL query:" & _
'                    vbNewLine & _
'                    """" & sqlStatement & """" & _
'                    vbNewLine & _
'                    "-------------------------------"
    End If

    'Close database recordset and connection
    '---------------------------------------
    rs.Close
    conn.Close
    'Memory cleanup
    Set rs = Nothing
    Set conn = Nothing

    SQLselect = True
End Function
