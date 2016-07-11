Option Compare Database
Public Function GetTagFromArg(ByVal OpenArgs As String, _
                               ByVal tag As String) As String
    Dim strArgument() As String
    strArgument = Split(OpenArgs, ":")
    Dim i As Integer
    For i = 0 To UBound(strArgument)
     If InStr(strArgument(i), tag) And _
                 InStr(strArgument(i), "=") > 0 Then
       GetTagFromArg = Mid$(strArgument(i), _
                        InStr(strArgument(i), "=") + 1)
       Exit Function
     End If
   Next
   GetTagFromArg = ""
End Function
Function IsNoData(vntCheckThis As Variant) As Boolean
Dim Result As Integer
    Result = False
On Error Resume Next
    If IsNull(vntCheckThis) Then Result = True
    If IsEmpty(vntCheckThis) Then Result = True
    If vntCheckThis = "" Then Result = True
    IsNoData = Result
End Function

Function fRefreshLinks(strNewPath As String, bolNewYear As Boolean) As Boolean
Dim strMsg As String, collTbls As Collection
Dim i As Integer, strDBPath As String, strTbl As String
Dim dbCurr As Database, dbLink As Database
Dim tdfLocal As TableDef
Dim varRet As Variant
Dim strCompPathLoc, strCompPathPasd, strCompPathSub1 As String
Dim bolCurrAppTbl As Boolean


Const cERR_USERCANCEL = vbObjectError + 1000
Const cERR_NOREMOTETABLE = vbObjectError + 2000

'current table belongs to current application
bolCurrAppTbl = False
'saves path to database passed into the function and path to Subdatabase
strCompPathPasd = Left(strNewPath, Len(strNewPath) - 8)
If bolNewYear Then
    strCompPathSub1 = strCompPathPasd
Else
    strCompPathSub1 = Left(Nz(Forms.frmChangeDB.lstMyCo.Column(4), ""), Abs(Len(Nz(Forms.frmChangeDB.lstMyCo.Column(4), "")) - 8))
End If
    On Local Error GoTo fRefreshLinks_Err

    If bolNewYear = False Then
        If MsgBox("Are you sure you want to Change Company?", _
                vbQuestion + vbYesNo, "Please confirm...") = vbNo Then Err.Raise cERR_USERCANCEL
    End If
    
    Set collTbls = fGetLinkedTables

    
    Set dbCurr = CurrentDb

    
    For i = collTbls.Count To 1 Step -1
        strDBPath = fParsePath(collTbls(i))
        strTbl = fParseTable(collTbls(i))
        varRet = SysCmd(acSysCmdSetStatus, "Now linking '" & strTbl & "'....")
        If Left$(strDBPath, 4) = "ODBC" Then
            'ODBC Tables
            'ODBC Tables handled separately
           ' Set tdfLocal = dbCurr.TableDefs(strTbl)
           ' With tdfLocal
           '     .Connect = pcCONNECT
           '     .RefreshLink
           '     collTbls.Remove (strTbl)
           ' End With
        Else
            If strNewPath <> vbNullString Then
                'Try this first
                strCompPathLoc = Left(strDBPath, Len(strDBPath) - 8)
                
                strCompPathLoc = Right(strCompPathLoc, Len(strCompPathLoc) - InStrRev(strCompPathLoc, "\"))
                strCompPathPasd = Right(strCompPathPasd, Len(strCompPathPasd) - InStrRev(strCompPathPasd, "\"))
                strCompPathSub1 = Right(strCompPathSub1, Len(strCompPathSub1) - InStrRev(strCompPathSub1, "\"))
                
                'chooses correct path for DB
'                Select Case strCompPathLoc
'                    Case strCompPathPasd: strDBPath = strNewPath
'                                          bolCurrAppTbl = True
'                    Case strCompPathSub1: strDBPath = Nz(Forms.frmChangeDB.lstMyCo.Column(4), strDBPath)
'                                          bolCurrAppTbl = False
'                    Case Else:            bolCurrAppTbl = False
'                End Select
            Else
                If Len(Dir(strDBPath)) = 0 Then
                    'File Doesn't Exist, call GetOpenFileName
                    strDBPath = fGetMDBName("'" & strDBPath & "' not found.")
                    If strDBPath = vbNullString Then
                        'user pressed cancel
                        Err.Raise cERR_USERCANCEL
                    End If
                End If
            End If

            'backend database exists
            'putting it here since we could have
            'tables from multiple sources
            Set dbLink = DBEngine(0).OpenDatabase(strNewPath)

            'check to see if the table is present in dbLink
            'strTbl = fParseTable(collTbls(i))
            If fIsRemoteTable(dbLink, strTbl) Then
                'everything's ok, reconnect
                bolCurrAppTbl = True
                strDBPath = strNewPath
                Set tdfLocal = dbCurr.TableDefs(strTbl)
                With tdfLocal
                    .connect = ";Database=" & strDBPath
                    .RefreshLink
                    collTbls.Remove (.name)
                End With
            Else
                bolCurrAppTbl = False
                Set dbLink = DBEngine(0).OpenDatabase(Nz(Forms.frmChangeDB.lstMyCo.Column(4), strDBPath))
                If fIsRemoteTable(dbLink, strTbl) Then
                    'everything's ok, reconnect
                    strDBPath = Nz(Forms.frmChangeDB.lstMyCo.Column(4), strDBPath)
                    Set tdfLocal = dbCurr.TableDefs(strTbl)
                    With tdfLocal
                        .connect = ";Database=" & strDBPath
                        .RefreshLink
                        collTbls.Remove (.name)
                    End With
                Else
                    Set dbLink = DBEngine(0).OpenDatabase(strDBPath)
                    If fIsRemoteTable(dbLink, strTbl) Then
                        'everything's ok, reconnect
                        Set tdfLocal = dbCurr.TableDefs(strTbl)
                        With tdfLocal
                            .connect = ";Database=" & strDBPath
                            .RefreshLink
                            collTbls.Remove (.name)
                        End With
                    Else
                        Err.Raise cERR_NOREMOTETABLE
                    End If
                End If
            End If
            
            If fsTableExists("prev" & tdfLocal.name) Then
               DoCmd.DeleteObject acTable, "prev" & tdfLocal.name
               collTbls.Remove ("prev" & tdfLocal.name)
               i = i - 1
            End If
            'do not look for other year if it's a system table
            If bolCurrAppTbl = True Then
                If fsDBPath(strDBPath, -1) <> Empty Then
                    Set tdfLocal = CurrentDb.CreateTableDef("prev" & strTbl)
                    With tdfLocal
                        .connect = ";Database=" & fsDBPath(strDBPath, -1)
                        .SourceTableName = strTbl
                        CurrentDb.TableDefs.Append tdfLocal
                    
'                        .RefreshLink

                    End With
                    Set tdfLocal = dbCurr.TableDefs(strTbl)
                End If
            
                If fsTableExists("next" & tdfLocal.name) Then
                   DoCmd.DeleteObject acTable, "next" & tdfLocal.name
                   collTbls.Remove ("next" & tdfLocal.name)
                   i = i - 1
                End If
                If fsDBPath(strDBPath, 1) <> Empty Then
                    Set tdfLocal = CurrentDb.CreateTableDef("next" & strTbl)
                    With tdfLocal
                        .connect = ";Database=" & fsDBPath(strDBPath, 1)
                        .SourceTableName = strTbl
                        CurrentDb.TableDefs.Append tdfLocal
'                       .RefreshLink

                    End With
                End If
            End If
        End If
        
move_next:
    Next
    fRefreshLinks = True
    varRet = SysCmd(acSysCmdClearStatus)
    MsgBox "All Access tables were successfully reconnected.", _
            vbInformation + vbOKOnly, _
            "Success"

fRefreshLinks_End:
    Set collTbls = Nothing
    Set tdfLocal = Nothing
    Set dbLink = Nothing
    Set dbCurr = Nothing
    Exit Function
fRefreshLinks_Err:
    fRefreshLinks = False
    Select Case Err
        Case 3059:

        Case cERR_USERCANCEL:
            MsgBox "Operation is canceled.", _
                    vbCritical + vbOKOnly, _
                    "Refreshing links."
            Resume fRefreshLinks_End
        Case cERR_NOREMOTETABLE:
            MsgBox "Table '" & strTbl & "' was not found in the specified company file" & _
                    vbCrLf & dbLink.name & ". Couldn't refresh links", _
                    vbCritical + vbOKOnly, _
                    "Refreshing links."
            Resume fRefreshLinks_End
        Case Else:
            strMsg = "Error Information..." & vbCrLf & vbCrLf
            strMsg = strMsg & "Function: fRefreshLinks" & vbCrLf
            strMsg = strMsg & "Description: " & Err.Description & vbCrLf
            strMsg = strMsg & "Error #: " & Format$(Err.Number) & vbCrLf
            MsgBox strMsg, vbOKOnly + vbCritical, "Error"
            Resume fRefreshLinks_End
    End Select
End Function

Function fIsRemoteTable(dbRemote As Database, strTbl As String) As Boolean
Dim tdf As TableDef
    On Error Resume Next
    Set tdf = dbRemote.TableDefs(strTbl)
    fIsRemoteTable = (Err = 0)
    Set tdf = Nothing
End Function

Function fGetMDBName(strIn As String) As String
'Calls GetOpenFileName dialog
''Dim strFilter As String

''    strFilter = ahtAddFilterItem(strFilter, _
                    "Access Database(*.mdb;*.mda;*.mde;*.mdw) ", _
                    "*.mdb; *.mda; *.mde; *.mdw")
''    strFilter = ahtAddFilterItem(strFilter, _
                    "All Files (*.*)", _
                    "*.*")

''    fGetMDBName = ahtCommonFileOpenSave(Filter:=strFilter, _
                                OpenFile:=True, _
                                DialogTitle:=strIn, _
                                Flags:=ahtOFN_HIDEREADONLY)
End Function

Function fGetLinkedTables() As Collection
'Returns all linked tables
    Dim collTables As New Collection
    Dim tdf As TableDef, Db As Database
    Set Db = CurrentDb
    Db.TableDefs.Refresh
    For Each tdf In Db.TableDefs
        With tdf
            If Len(.connect) > 0 Then
                If Left$(.connect, 4) = "ODBC" Then
                '    collTables.Add Item:=.Name & ";" & .Connect, KEY:=.Name
                'ODBC Reconnect handled separately
                Else
                    collTables.Add item:=.name & .connect, Key:=.name
                End If
            End If
        End With
    Next
    Set fGetLinkedTables = collTables
    Set collTables = Nothing
    Set tdf = Nothing
    Set Db = Nothing
End Function
Function FnIme() As String
    Dim rs As Recordset
    Dim strSQL As String
    Dim i As Integer

    strSQL = "SELECT ctrName, yearID FROM tblControl;"

    Set rs = CurrentDb.OpenRecordset(strSQL)

    FnIme = rs.Fields("ctrName") & " ( " & rs.Fields("yearID") & " )"
    rs.Close
End Function
Function fParsePath(strIn As String) As String
    If Left$(strIn, 4) <> "ODBC" Then
        fParsePath = Right(strIn, Len(strIn) _
                        - (InStr(1, strIn, "DATABASE=") + 8))
    Else
        fParsePath = strIn
    End If
End Function

Function fParseTable(strIn As String) As String
    fParseTable = Left$(strIn, InStr(1, strIn, ";") - 1)
End Function
Function FnCOF(Optional FormName As String)
Dim i As Integer
Dim currentForm As Form
On Error GoTo Err_End
    'For i = 1 To Forms.Count
        'Set currentForm = Screen.ActiveForm
    For Each currentForm In Application.Forms
        If currentForm.Modal = False Then
             
             If currentForm.NewRecord = True Then
                Select Case MsgBox("Save Current Record and Exit? YES or NO?" _
                    , vbYesNo + vbDefaultButton1, "Warning!")
                Case vbYes
                    DoCmd.RunCommand acCmdSaveRecord
                Case vbNo
                    Exit Function
                End Select
             Else
             End If
             
             Select Case currentForm.CurrentRecord
             Case 0, 1
                GoTo Prazna
             Case Else
                DoCmd.GoToRecord acActiveDataObject, , acPrevious
             End Select
        
        Else
        End If
Prazna:
        DoCmd.Close acForm, currentForm.name
    'Next i
    Next currentForm
If FormName <> "" Then
    DoCmd.OpenForm FormName
Else
End If
Exit Function
Err_End:
If Err.Number = 2455 Then GoTo Prazna
MsgBox "Incorrect Input in Current Form!", vbExclamation
Debug.Print Err.Number
End Function

Public Function SplitRef(ByVal RefCode As String, ByVal level As Byte) As String
    Select Case level
        Case 1
            SplitRef = Left(RefCode, 2)
        Case 2
            SplitRef = Right(Left(RefCode, 4), 2)
        Case 3
            SplitRef = Right(Left(RefCode, 6), 2)
        Case 4
            SplitRef = Right(Left(RefCode, 8), 2)
        Case 5
            SplitRef = Right(RefCode, 3)
    End Select
    
End Function

Function GetDBFName(intOrder As Integer) As String
'Order - -1 previous year database, 0 current year database, 1 next year database
Dim rs As Recordset

strSQL = "SELECT sysMyCo.smycoPath, sysMyCo.smycoYear " & _
         "FROM tblControl INNER JOIN sysMyCo ON tblControl.ctrDBID = sysMyCo.smycoID;"

Set rs = CurrentDb.OpenRecordset(strSQL)
rs.MoveFirst
GetDBFName = Left(rs.Fields(0), Len(rs.Fields(0)) - 9) & "_" & (rs.Fields(1) + intOrder) & ".mdb"
'GetDBFName = Right(GetDBFName, Len(GetDBFName) - InStrRev(GetDBFName, "\"))

End Function
Function fsDBPath(strCurrPath As String, intOrder As Integer) As String
'Order - -1 previous year database, 0 current year database, 1 next year database
Dim intYear As Integer
Dim strNewPath As String

intYear = CInt(Left(Right(strCurrPath, 8), 4))
strNewPath = Left(strCurrPath, Len(strCurrPath) - 8) & CStr(intYear + intOrder) & ".mdb"

If Dir(strNewPath) <> "" Then
    fsDBPath = strNewPath
Else
    fsDBPath = Empty
End If
End Function

Function CreateNewMDBFile() As String

    Dim oAcc As Access.Application
    Dim ws As Workspace
    Dim Db As Database
    Dim NewDBPath As String
    Dim CurrDBPath As String

    'Get default Workspace
    Set ws = DBEngine.Workspaces(0)

    'Path and file name for new mdb file
    NewDBPath = GetDBFName(1)

    'Make sure there isn't already a file with the name of the new database
    If Dir(NewDBPath) <> "" Then Kill NewDBPath

    'Create a new mdb file
    Set Db = ws.CreateDatabase(NewDBPath, dbLangGeneral)
    Db.Close
    
    'Path and file name for current mdb file
    CurrDBPath = GetDBFName(0)

    Set oAcc = New Access.Application
    Set Db = oAcc.DBEngine.OpenDatabase(CurrDBPath, _
    False, False)
    oAcc.OpenCurrentDatabase CurrDBPath

    'For lookup tables, export both table definition and data to new mdb file
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblAccount", "tblAccount", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblBank", "tblBank", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblBankAccount", "tblBankAccount", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblChart", "tblChart", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblControl", "tblControl", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblCurrency", "tblCurrency", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblDescription", "tblDescription", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblLevel", "tblLevel", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblPeriod", "tblPeriod", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblPeriodSel", "tblPeriodSel", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblReference", "tblReference", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblTransaction", "tblTransaction", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblTransactionSub", "tblTransactionSub", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblVat", "tblVat", False
    oAcc.DoCmd.TransferDatabase acExport, "Microsoft Access", NewDBPath, acTable, "tblYear", "tblYear", False

    Db.Close
    oAcc.CloseCurrentDatabase
    oAcc.Quit acExit
    Set oAcc = Nothing
    Set Db = Nothing
    CreateNewMDBFile = NewDBPath
    
End Function

Function fsTableExists(tableName As String) As Boolean

Dim strTableNameCheck
On Error GoTo ErrorCode

'try to assign tablename value
strTableNameCheck = CurrentDb.TableDefs(tableName)

'If no error and we get to this line, true
fsTableExists = True

ExitCode:
    On Error Resume Next
    Exit Function

ErrorCode:
    Select Case Err.Number
        Case 3265  'Item not found in this collection
            fsTableExists = False
            Resume ExitCode
        Case Else
            MsgBox FnIsErr(Err.Number), vbExclamation
            'Debug.Print "Error " & Err.number & ": " & Err.Description & "hlfUtils.TableExists"
            Resume ExitCode
    End Select

End Function

Function fsClosedYear() As Boolean
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("SELECT ctrClosed FROM tblControl")
fsClosedYear = rs.Fields(0)
rs.Close
Set rs = Nothing
End Function
Function fsPrevYearExist() As Boolean
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("SELECT smycoYear FROM sysMyCo WHERE smyCoYear = " & DLookup("yearID", "tblControl") - 1)
If rs.RecordCount > 0 Then
    fsPrevYearExist = True
Else
    fsPrevYearExist = False
End If
rs.Close
Set rs = Nothing
End Function
Function fsNextYearHasChanged() As Boolean
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("SELECT ctrHasChanged FROM nexttblControl")
fsNextYearHasChanged = rs.Fields(0)
rs.Close
Set rs = Nothing
End Function

Function fsMoreID(LastID As Long) As Boolean
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("SELECT Max(trnsID) FROM tblTransactionSub")
If LastID < rs.Fields(0) Then
    fsMoreID = True
Else
    fsMoreID = False
End If
rs.Close
Set rs = Nothing
End Function
'updates opening balances of next year based on transactions in current db year
Sub sUpdtBalances()
    Dim rs As Recordset
    Dim rsCheck As Recordset
    Dim strList As String
    
    DoCmd.SetWarnings False
    strSQL = "SELECT Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount]))+tblAccount.accOpenBalance AS Bal, tblAccount.accNo, tblAccount.caoID, tblAccount.accName " & _
             "FROM tblAccount INNER JOIN (tblTransactionSub INNER JOIN tblTransaction ON tblTransactionSub.trnID = tblTransaction.trnID) ON tblAccount.accID = tblTransactionSub.accID " & _
             "WHERE (((tblTransaction.trnYear) In (SELECT yearID FROM tblControl)) AND ((tblAccount.accType)=2)) " & _
             "GROUP BY tblAccount.accNo, tblAccount.accOpenBalance, tblAccount.caoID, tblAccount.accName;"
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    'Balance sheets
    rs.MoveFirst
    Do Until rs.EOF
        dblBalance = CDbl(rs.Fields("Bal").Value)
        strSQL = "SELECT nexttblAccount.caoID, nexttblAccount.accName, nexttblAccount.accOpenBalance " & _
                 "FROM nexttblAccount " & _
                 "WHERE nexttblAccount.accNo = '" & rs.Fields("accNo") & "'"
        Set rsCheck = CurrentDb.OpenRecordset(strSQL)
        
        If rsCheck.RecordCount > 0 Then
            If (rsCheck.Fields("accName") = rs.Fields("accName")) _
            And (rsCheck.Fields("caoID") = rs.Fields("caoID")) Then
                strSQL = "UPDATE nexttblAccount " & _
                         "SET accOpenBalance = " & Replace(CStr(rs.Fields("Bal")), ",", ".") & " " & _
                         "WHERE accNo = '" & rs.Fields("accNo") & "'"
                FnLog (strSQL)
                DoCmd.RunSQL (strSQL)
            Else
                strList = strList & "Acc No: " & rs.Fields("accNo") & Chr(9) & _
                          "O/N Acc Name: " & rs.Fields("accName") & " / " & rsCheck.Fields("accName") & Chr(9) & _
                          "O/N Ref ID: " & rs.Fields("caoID") & " / " & rsCheck.Fields("caoID") & Chr(9) & _
                          "N bal: " & rs.Fields("Bal") & Chr(13)
            End If
        Else
            strList = strList & "Acc No: " & rs.Fields("accNo") & " does not exist! " & Chr(9) & _
                      "Bal: " & rs.Fields("Bal") & Chr(13)
        End If
        rs.MoveNext
    Loop
    
    'Profit and Lose
'    strSQL = "SELECT accID FROM tblAccount WHERE accName"
    strSQL = "SELECT Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount])) AS Bal, tblControl.ctrLYPLA " & _
             "FROM tblControl, (tblAccount INNER JOIN (tblTransactionSub INNER JOIN tblTransaction ON tblTransactionSub.trnID = tblTransaction.trnID) ON tblAccount.accID = tblTransactionSub.accID) INNER JOIN nexttblAccount ON tblAccount.accID = nexttblAccount.accID " & _
             "WHERE (((tblTransaction.trnYear) In (SELECT yearID FROM tblControl)) AND ((tblAccount.accType)=1)) " & _
             "GROUP BY tblControl.ctrLYPLA;"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount Then
        rs.MoveFirst
        
        strSQL = "SELECT nexttblAccount.accID " & _
                 "FROM nexttblAccount " & _
                 "WHERE nexttblAccount.accID = " & rs.Fields(1)
        Set rsCheck = CurrentDb.OpenRecordset(strSQL)
           
        If rsCheck.RecordCount = 0 Then
            strList = strList & "Profit&Loss Acc ID: " & rs.Fields(1) & " does not exist! " & Chr(9) & _
                     "Bal: " & rs.Fields(0) & Chr(13)
        Else
            strSQL = "UPDATE nexttblAccount " & _
                     "SET accOpenBalance = " & rs.Fields(0) & _
                     " WHERE accID = " & rs.Fields(1)
                 
            FnLog (strSQL)
            DoCmd.RunSQL (strSQL)
        End If
    End If
    DoCmd.SetWarnings True
        
    strSQL = "SELECT sysMyCo.smycoPath " & _
             "FROM tblControl INNER JOIN sysMyCo ON tblControl.ctrDBID = sysMyCo.smycoID " & _
             "GROUP BY sysMyCo.smycoPath; "

    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    Open Left(rs.Fields(0), InStrRev(rs.Fields(0), "\")) & "Log_lleger.txt" For Output As #1
    Write #1, strList
    Close #1

    rs.Close
    Set rs = Nothing
    rsCheck.Close
    Set rsCheck = Nothing

End Sub
'updates opening balances of current db year based on transactions in previous year
Sub sUpdtBalancesCur()
    Dim rs As Recordset
    Dim rsCheck As Recordset
    Dim strList As String
    
    DoCmd.SetWarnings False
    
    'initialize the to zero balance
    strSQL = "UPDATE tblAccount SET tblAccount.accOpenBalance = 0;"
    DoCmd.RunSQL (strSQL)
    
    strSQL = "SELECT Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount]))+prevtblAccount.accOpenBalance AS Bal, prevtblAccount.accNo, prevtblAccount.caoID, prevtblAccount.accName " & _
             "FROM prevtblAccount INNER JOIN (prevtblTransactionSub INNER JOIN prevtblTransaction ON prevtblTransactionSub.trnID = prevtblTransaction.trnID) ON prevtblAccount.accID = prevtblTransactionSub.accID " & _
             "WHERE (((prevtblTransaction.trnYear) In (SELECT yearID FROM prevtblControl)) AND ((prevtblAccount.accType)=2)) " & _
             "GROUP BY prevtblAccount.accNo, prevtblAccount.accOpenBalance, prevtblAccount.caoID, prevtblAccount.accName;"
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    'Balance sheets
    rs.MoveFirst
    Do Until rs.EOF
        dblBalance = CDbl(rs.Fields("Bal").Value)
        strSQL = "SELECT tblAccount.caoID, tblAccount.accName, tblAccount.accOpenBalance " & _
                 "FROM tblAccount " & _
                 "WHERE tblAccount.accNo = '" & rs.Fields("accNo") & "'"
        Set rsCheck = CurrentDb.OpenRecordset(strSQL)
        
        If rsCheck.RecordCount > 0 Then
            If (rsCheck.Fields("accName") = rs.Fields("accName")) _
            And (rsCheck.Fields("caoID") = rs.Fields("caoID")) Then
                strSQL = "UPDATE tblAccount " & _
                         "SET accOpenBalance = " & Replace(CStr(rs.Fields("Bal")), ",", ".") & " " & _
                         "WHERE accNo = '" & rs.Fields("accNo") & "'"
                FnLog (strSQL)
                DoCmd.RunSQL (strSQL)
            Else
                strList = strList & "Acc No: " & rs.Fields("accNo") & Chr(9) & _
                          "O/N Acc Name: " & rs.Fields("accName") & " / " & rsCheck.Fields("accName") & Chr(9) & _
                          "O/N Ref ID: " & rs.Fields("caoID") & " / " & rsCheck.Fields("caoID") & Chr(9) & _
                          "N bal: " & rs.Fields("Bal") & Chr(13)
            End If
        Else
            strList = strList & "Acc No: " & rs.Fields("accNo") & " does not exist! " & Chr(9) & _
                      "Bal: " & rs.Fields("Bal") & Chr(13)
        End If
        rs.MoveNext
    Loop
    
    'Profit and Lose
    'was taking the account from previous control
'    strSQL = "SELECT accID FROM tblAccount WHERE accName"
'    strSQL = "SELECT Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount])) AS Bal, prevtblControl.ctrLYPLA " & _
'             "FROM prevtblControl, (prevtblAccount INNER JOIN (prevtblTransactionSub INNER JOIN prevtblTransaction ON prevtblTransactionSub.trnID = prevtblTransaction.trnID) ON prevtblAccount.accID = prevtblTransactionSub.accID) INNER JOIN tblAccount ON prevtblAccount.accID = tblAccount.accID " & _
'             "WHERE (((prevtblTransaction.trnYear) In (SELECT yearID FROM prevtblControl)) AND ((prevtblAccount.accType)=1)) " & _
'             "GROUP BY prevtblControl.ctrLYPLA;"
 'was taking the profit &loss account  control
   strSQL = "SELECT Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount])) AS Bal, tblControl.ctrLYPLA " & _
            "FROM tblControl, (prevtblAccount INNER JOIN (prevtblTransactionSub INNER JOIN prevtblTransaction ON prevtblTransactionSub.trnID = prevtblTransaction.trnID) ON prevtblAccount.accID = prevtblTransactionSub.accID) INNER JOIN tblAccount ON prevtblAccount.accID = tblAccount.accID " & _
            "WHERE (((prevtblTransaction.trnYear) In (SELECT yearID FROM prevtblControl)) AND ((prevtblAccount.accType)=1))" & _
            "GROUP BY tblControl.ctrLYPLA;"
             
             
             
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount Then
        rs.MoveFirst
    
        strSQL = "SELECT tblAccount.accID " & _
                 "FROM tblAccount " & _
                 "WHERE tblAccount.accID = " & rs.Fields(1)
        Set rsCheck = CurrentDb.OpenRecordset(strSQL)
           
        If rsCheck.RecordCount = 0 Then
            strList = strList & "Profit&Loss Acc ID: " & rs.Fields(1) & " does not exist! " & Chr(9) & _
                     "Bal: " & rs.Fields(0) & Chr(13)
        Else
            strSQL = "UPDATE tblAccount " & _
                     "SET accOpenBalance = " & rs.Fields(0) & _
                     " WHERE accID = " & rs.Fields(1)
            FnLog (strSQL)
            DoCmd.RunSQL (strSQL)
        End If
    End If
    DoCmd.SetWarnings True
    'log errors
    strSQL = "SELECT sysMyCo.smycoPath " & _
             "FROM tblControl INNER JOIN sysMyCo ON tblControl.ctrDBID = sysMyCo.smycoID " & _
             "GROUP BY sysMyCo.smycoPath; "

    Set rs = CurrentDb.OpenRecordset(strSQL)
    'strList = Left(rs.Fields(0), InStrRev(rs.Fields(0), "\")) & "Log_lleger.txt"
    
    Open Left(rs.Fields(0), InStrRev(rs.Fields(0), "\")) & "Log_lleger.txt" For Output As #1
    'Open strList For Output As #1
    Write #1, strList
    Close #1

    rs.Close
    Set rs = Nothing
    rsCheck.Close
    Set rsCheck = Nothing

End Sub

Sub sUpdateAccount(accNo As Long)
Dim rs As Recordset
Dim strSQL As String

strSQL = "SELECT accID FROM nexttblAccount WHERE accNo = '" & accNo & "'"

Set rs = CurrentDb.OpenRecordset(strSQL)
If rs.RecordCount > 0 Then
    strSQL = "UPDATE nexttblAccount " & _
             "SET caoID = " & Forms.frmAccount.Controls("caoID").Value & ", " & _
             "    accNo = " & Forms.frmAccount.Controls("accNo").Value & ", " & _
             "    accName = '" & Forms.frmAccount.Controls("accName").Value & "', " & _
             "    accStatus = " & Forms.frmAccount.Controls("accStatus").Value & ", " & _
             "    accIsVat = " & Forms.frmAccount.Controls("accIsVat").Value & ", " & _
             "    accSign = '" & Forms.frmAccount.Controls("accSign").Value & "', " & _
             "    accType = " & Forms.frmAccount.Controls("accType").Value & ", " & _
             "    accVatable = " & Forms.frmAccount.Controls("accVatable").Value & " " & _
             "WHERE accNo = '" & accNo & "'"
Else
    strSQL = "INSERT INTO nexttblAccount(caoID, accNo, accName, accStatus, accIsVat, accSign, accType, accOpenBalance, accVatable) " & _
                    "VALUES(" & Forms.frmAccount.Controls("caoID").Value & ", " _
                              & Forms.frmAccount.Controls("accNo").Value & ", '" _
                              & Forms.frmAccount.Controls("accName").Value & "', " _
                              & Forms.frmAccount.Controls("accStatus").Value & ", " _
                              & Forms.frmAccount.Controls("accIsVat").Value & ", '" _
                              & Forms.frmAccount.Controls("accSign").Value & "', " _
                              & Forms.frmAccount.Controls("accType").Value & ", " _
                              & Forms.frmAccount.Controls("accOpBal").Value & ", " _
                              & Forms.frmAccount.Controls("accVatable").Value & ")"
End If
DoCmd.SetWarnings False
FnLog (strSQL)
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True
rs.Close
Set rs = Nothing
Forms.frmAccount.Controls("HasChanged").Value = False
End Sub
Sub ChartAccType(coaRef1 As Integer, coaRef2 As Integer, coaRef3 As Integer, coaRef4 As Integer)

Dim rs As Recordset
Dim lvl1, lvl2, lvl3, lvl4 As Byte
lvl1 = 0
lvl2 = 0
lvl3 = 0
lvl4 = 0

strSQL = "SELECT tblAccount.accType, 1 As Lvl " & _
          "FROM tblChart INNER JOIN tblAccount ON tblChart.coaID = tblAccount.caoID " & _
          "WHERE (((tblChart.lvlID) = 5) And ((tblChart.coaRef1) = " & coaRef1 & _
          ")) GROUP BY tblAccount.accType" & _
          " UNION ALL " & _
          "SELECT tblAccount.accType, 2 As Lvl " & _
          "FROM tblChart INNER JOIN tblAccount ON tblChart.coaID = tblAccount.caoID " & _
          "WHERE (((tblChart.lvlID) = 5) AND ((tblChart.coaRef1) = " & coaRef1 & ") AND ((tblChart.coaRef2) = " & coaRef2 & _
          ")) GROUP BY tblAccount.accType" & _
          " UNION ALL " & _
          " SELECT tblAccount.accType, 3 As Lvl " & _
          "FROM tblChart INNER JOIN tblAccount ON tblChart.coaID = tblAccount.caoID " & _
          "WHERE (((tblChart.lvlID) = 5) " & _
          "AND ((tblChart.coaRef1) = " & coaRef1 & ") AND ((tblChart.coaRef2) = " & coaRef2 & ") AND ((tblChart.coaRef3) = " & coaRef3 & _
          ")) GROUP BY tblAccount.accType" & _
          " UNION ALL " & _
          " SELECT tblAccount.accType, 4 As Lvl " & _
          "FROM tblChart INNER JOIN tblAccount ON tblChart.coaID = tblAccount.caoID " & _
          "WHERE (((tblChart.lvlID) = 5) " & _
          "AND ((tblChart.coaRef1) = " & coaRef1 & ") AND ((tblChart.coaRef2) = " & coaRef2 & ") AND ((tblChart.coaRef3) = " & coaRef3 & ") AND ((tblChart.coaRef4) = " & coaRef4 & _
          ")) GROUP BY tblAccount.accType"

Set rs = CurrentDb.OpenRecordset(strSQL)

If rs.RecordCount <> 0 Then
    rs.MoveFirst

    Do Until rs.EOF
        Select Case rs.Fields("Lvl")
        Case 1
            lvl1 = lvl1 + rs.Fields("accType")
        Case 2
            lvl2 = lvl2 + rs.Fields("accType")
        Case 3
            lvl3 = lvl3 + rs.Fields("accType")
        Case 4
            lvl4 = lvl4 + rs.Fields("accType")
        
        End Select
        rs.MoveNext
    Loop

    DoCmd.SetWarnings False
    strSQL = "UPDATE tblChart SET coaAccType = " & lvl1 & " WHERE lvlID = 1 AND coaRef1 = " & coaRef1
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    strSQL = "UPDATE tblChart SET coaAccType = " & lvl2 & " WHERE lvlID = 2 AND coaRef1 = " & coaRef1 & " AND coaRef2 = " & coaRef2
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    strSQL = "UPDATE tblChart SET coaAccType = " & lvl3 & " WHERE lvlID = 3 AND coaRef1 = " & coaRef1 & " AND coaRef2 = " & coaRef2 & " AND coaRef3 = " & coaRef3
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    strSQL = "UPDATE tblChart SET coaAccType = " & lvl4 & " WHERE lvlID = 4 AND coaRef1 = " & coaRef1 & " AND coaRef2 = " & coaRef2 & " AND coaRef3 = " & coaRef3 & " AND coaRef4 = " & coaRef4
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings True

End If

End Sub
Public Function SetNumeric(Var As Variant) As Double

If IsNumeric(Var) = True Then
    SetNumeric = CDbl(Var)
Else
    SetNumeric = 0
End If

End Function
Function RoundDG(ByVal Number As Double) As Double
    If (Number - (Int(Number * 100) / 100)) < 0.005 Then
        RoundDG = (Int(Number * 100) / 100)
    Else
        RoundDG = (Int((Number + 0.005) * 100) / 100)
    End If
    'RoundDigit = (Int(Number * 100) / 100)
End Function

Function CopyLastTrnsField(fldName As String, trnID As Integer, Optional varReturn As Variant = "") As Variant
    Dim strSQL As String
    Dim rs As Recordset
    
    strSQL = "SELECT * FROM tblTransactionSub WHERE tblTransactionSub.trnID = " & trnID & " Order By trnsID DESC"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        CopyLastTrnsField = rs.Fields(fldName)
    Else
        CopyLastTrnsField = varReturn
    End If
        
End Function