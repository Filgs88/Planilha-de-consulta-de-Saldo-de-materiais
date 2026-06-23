Sub CriarQueryTable()
    Dim strConn As String
    Dim strSQL As String
    Dim qt As QueryTable

    strConn = "ODBC;Driver={DRIVER};Dbq=IP:PORT/SERVER-NAME;Uid=USER;Pwd=PASSWORD;"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arquivo = fso.OpenTextFile(".\sql_estoque_de_materiais.sql", 1)
    strSQL = arquivo.ReadAll

    Planilha1.Cells.Clear

    Set qt = Planilha1.QueryTables.Add(Connection:=strConn, Destination:=Planilha1.Range("A1"), Sql:=strSQL)

    With qt
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .Refresh
    End With
End Sub
