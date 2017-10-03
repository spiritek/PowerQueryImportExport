Attribute VB_Name = "PowerQueryImportExport"
Function DoesQueryExist(ByVal queryName As String) As Boolean
    ' This function is from script here https://gallery.technet.microsoft.com/office/VBA-to-automate-Power-956a52d1
    ' by Gil Raviv https://social.technet.microsoft.com/profile/gil%20raviv/
    ' Helper function to check if a query with the given name already exists
    Dim qry As WorkbookQuery
     
    If (ActiveWorkbook.Queries.Count = 0) Then
        DoesQueryExist = False
        Exit Function
    End If
     
    For Each qry In ActiveWorkbook.Queries
        If (qry.Name = queryName) Then
            DoesQueryExist = True
            Exit Function
        End If
    Next
    DoesQueryExist = False
End Function
Sub CleanSheet(ByVal sheetName As String)
    ' Helper function to try to delete the worksheet if exists
    On Error Resume Next
    ActiveWorkbook.Sheets(sheetName).Delete
End Sub


Sub LoadToWorksheetOnly(query As WorkbookQuery, currentSheet As Worksheet)
    ' This function is from script here https://gallery.technet.microsoft.com/office/VBA-to-automate-Power-956a52d1
    ' by Gil Raviv https://social.technet.microsoft.com/profile/gil%20raviv/
    '
    ' The usual VBA code to create ListObject with a Query Table
    ' The interface is not new, but looks how simple is the conneciton string of Power Query:
    ' "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name
     
    With currentSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdDefault
        .CommandText = Array("SELECT * FROM [" & query.Name & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .Refresh BackgroundQuery:=False
    End With
     
End Sub
 
Sub LoadToWorksheetAndModel(query As WorkbookQuery, currentSheet As Worksheet)
    ' This function is from script here https://gallery.technet.microsoft.com/office/VBA-to-automate-Power-956a52d1
    ' by Gil Raviv https://social.technet.microsoft.com/profile/gil%20raviv/
    
    ' Let's load the query to the Data Model
    LoadToDataModel query
     
    ' Now we can load the data to the worksheet
    With currentSheet.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook. _
        Connections("Query - " & query.Name), Destination:=Range("$A$1")).TableObject
        .RowNumbers = False
        .PreserveFormatting = True
        .PreserveColumnInfo = False
        .AdjustColumnWidth = True
        .RefreshStyle = 1
        .ListObject.DisplayName = Replace(query.Name, " ", "_") & "_ListObject"
        .Refresh
    End With
End Sub
 
Sub LoadToDataModel(query As WorkbookQuery)
    ' This function is from script here https://gallery.technet.microsoft.com/office/VBA-to-automate-Power-956a52d1
    ' by Gil Raviv https://social.technet.microsoft.com/profile/gil%20raviv/
    ' This code loads the query to the Data Model
    ActiveWorkbook.Connections.Add2 "Query - " & query.Name, _
        "Connection to the '" & query.Name & "' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name _
        , """" & query.Name & """", 6, True, False
 
End Sub


Sub ExportPowerQueryQueries()
Attribute ExportPowerQueryQueries.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
    Dim oFs As Object
    Dim oFile As Object
    Dim qPath As String
    Dim numQry As Integer
    confirmation = MsgBox("This will replace any existing power query files", vbOKCancel, "Are you sure?")
    If confirmation = 1 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            qPath = .SelectedItems(1)
        End With
        
        qPath = qPath & "\"
        
        numQry = ActiveWorkbook.Queries.Count
        For i = 1 To numQry
            qName = ActiveWorkbook.Queries.Item(i).Name
            qPathFile = qPath & qName & ".pq"
            qContent = ActiveWorkbook.Queries.Item(i).Formula
            Set oFs = CreateObject("Scripting.FileSYstemObject")
            Set oFile = oFs.CreateTextFile(qPathFile)
            oFile.WriteLine qContent
            oFile.Close
        Next
        Set oFs = Nothing
        Set oFile = Nothing
        Set oFolder = Nothing
    Else
        MsgBox "Export Cancelled"
    End If
End Sub

Sub ImportPowerQueryQueries()
'
'
    Dim oFs As Object
    Dim oFile As Object
    Dim oFolder As Object
    Dim qPath As String
    Dim qry As WorkbookQuery
    confirmation = MsgBox("This will replace any existing power query files", vbOKCancel, "Are you sure?")
    If confirmation = 1 Then
        
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            qPath = .SelectedItems(1)
        End With

        qPath = qPath & "\"
        Set oFs = CreateObject("Scripting.FileSYstemObject")
        Set oFolder = oFs.GetFolder(qPath)
        Set cFiles = oFolder.Files
        For Each oFile In cFiles
            If Right(oFile.Name, 3) = ".pq" Then
                oFileName = oFile.Name
                Set oReadFile = oFs.OpenTextFile(oFile, 1)
                qContent = oReadFile.ReadAll
                qName = Left(oFileName, Len(oFileName) - 3)
                If DoesQueryExist(qName) Then
                    ActiveWorkbook.Queries.Item(qName).Delete
                    CleanSheet (qName)
                End If
                Set qry = ActiveWorkbook.Queries.Add(qName, qContent)
                LoadToDataModel qry
                oReadFile.Close
            End If
        Next
        Set oFs = Nothing
        Set oFile = Nothing
        Set oFolder = Nothing
    Else
        MsgBox "Import Cancelled"
    End If
End Sub


