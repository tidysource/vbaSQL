Attribute VB_Name = "VBA_require_SQL"
Option Explicit

'Run table definistions
'----------------------
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
