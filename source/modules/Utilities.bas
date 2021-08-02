Option Compare Database

Public Sub ExecuteSQLScript(SQLFilepath As String, AccessFile As String)

    Dim Querytext As String
    Dim Query As DAO.QueryDef
    Dim vSQL As Variant
    Dim vSQLs As Variant
    Dim Access As Access.Application
    
    Set Access = GetObject(AccessFile)
    
    'queryfile laden
    Querytext = LoadTextFile(SQLFilepath)
    'query definieren
    vSQL = Split(Querytext, ";")
    Access.DoCmd.SetWarnings False
    On Error Resume Next
    For Each vSQLs In vSQL
        Set Query = Access.CurrentDb.CreateQueryDef("SQLScript", vSQLs)
        'query ausführen
        Access.DoCmd.OpenQuery ("SQLScript")
        'query löschen
        Access.CurrentDb.QueryDefs.Delete Query.name
    Next
    DoCmd.SetWarnings True
End Sub

Public Function LoadTextFile(filePath As String) As String
    Dim iFile As Integer: iFile = FreeFile
    Open filePath For Input As #iFile
    LoadTextFile = Input(LOF(iFile), iFile)
    Close iFile
End Function
Public Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
      MsgBox ("Datei gelöscht")
   End If
End Sub

Public Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function
 Public Sub CopyDB(NewFileName As String, SourceFile As String)
    Dim external_db As Object
    Dim sDBsource As String
    Dim sDBdest As String

    sDBsource = SourceFile
    sDBdest = NewFileName
    
    Set external_db = CreateObject("Scripting.FileSystemObject")
    external_db.CopyFile sDBsource, sDBdest, True
    Set external_db = Nothing
End Sub