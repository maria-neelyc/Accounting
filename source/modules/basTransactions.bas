Option Compare Database

Function fsAmount(lngID As Long) As Double
Dim strSQL As String
strSQL = "SELECT Sum(IIf([trnsSign]='C',[trnsAmount],-1*[trnsAmount])) AS tAmount " _
         & "FROM tblTransactionSub " _
         & "WHERE (((tblTransactionSub.trnID)=" & lngID & "));"

Dim Db As Database
Dim rs As DAO.Recordset
Set Db = CurrentDb()
Set rs = Db.OpenRecordset(strSQL)
rs.MoveFirst
'rs ()
If IsNoData(rs(0).Value) = True Then
    fsAmount = 0
Else
    fsAmount = rs(0).Value
End If

rs.Close
Db.Close
End Function

Function fsFAmount(lngID As Long) As Double
Dim strSQL As String
strSQL = "SELECT Sum(IIf([trnsSign]='C',[trnsFAmount],-1*[trnsFAmount])) AS tAmount " _
         & "FROM tblTransactionSub " _
         & "WHERE (((tblTransactionSub.trnID)=" & lngID & "));"

Dim Db As Database
Dim rs As DAO.Recordset
Set Db = CurrentDb()
Set rs = Db.OpenRecordset(strSQL)
rs.MoveFirst
'rs ()
If IsNoData(rs(0).Value) = True Then
    fsFAmount = 0
Else
    fsFAmount = rs(0).Value
End If

rs.Close
Db.Close
End Function

Function findVAT(dtEntryDate As Date) As Double
On Error GoTo Err_End
Dim strSQL As String
Dim strDate As String
        strDate = CStr(Format(dtEntryDate, "mm/dd/yyyy"))
        
        strSQL = "SELECT tblVat.vatID, tblVat.vatRate, tblVat.vatStartDate, tblVat.vatEndDate " & _
                 "FROM tblVat WHERE (((tblVat.vatStartDate)<=#" & _
                 strDate & "#) AND ((tblVat.vatEndDate)>=#" & strDate & "#));"
        Debug.Print strSQL
        Dim Db As Database
        Dim rs As DAO.Recordset
        Set Db = CurrentDb()
        Set rs = Db.OpenRecordset(strSQL)
        rs.MoveFirst
        'rs ()
        If IsNoData(rs(0).Value) = True Then
            findVAT = 0
        Else
            findVAT = rs("vatRate").Value
        End If

    rs.Close
    Db.Close
Exit Function
Err_End:
    findVAT = 0
    rs.Close
    Db.Close
End Function

Function fsCR_DB(lngID As Long, Field As String) As Double
Dim strSQL As String
strSQL = "SELECT Sum(" & Field & ") AS tAmount " _
         & "FROM tblTransactionSub " _
         & "WHERE (((tblTransactionSub.trnID)=" & lngID & "));"

Dim Db As Database
Dim rs As DAO.Recordset
Set Db = CurrentDb()
Set rs = Db.OpenRecordset(strSQL)
rs.MoveFirst
'rs ()
If IsNoData(rs(0).Value) = True Then
    fsCR_DB = 0
Else
    fsCR_DB = rs(0).Value
End If

rs.Close
Db.Close
End Function

Function fsSequence(intSRange As Integer, intERange As Integer) As Boolean
Dim strSQL As String
Dim i As Long

i = DLookup("Max(trnsSequence)", "tblTransactionSub") + 1

If intSRange = 0 And intERange = 0 Then
strSQL = "SELECT tblTransaction.trnPageCounter, tblTransaction.persID, tblTransactionSub.trnsDate, tblTransactionSub.trnsSequence " & _
         "FROM tblTransaction INNER JOIN tblTransactionSub ON tblTransaction.trnID = tblTransactionSub.trnID " & _
         "WHERE (((tblTransaction.trnLock) = True) And ((tblTransactionSub.trnsSequence) = 0)) " & _
         "ORDER BY tblTransaction.persID, tblTransactionSub.trnsDate, tblTransaction.trnPageCounter;"
Else
strSQL = "SELECT tblTransaction.trnPageCounter, tblTransaction.persID, tblTransactionSub.trnsDate, tblTransactionSub.trnsSequence " & _
         "FROM tblTransaction INNER JOIN tblTransactionSub ON tblTransaction.trnID = tblTransactionSub.trnID " & _
         "WHERE tblTransaction.trnLock = True AND tblTransaction.trnPageCounter >= " & intSRange & _
         " AND tblTransaction.trnPageCounter <= " & intERange & _
         " AND tblTransactionSub.trnsSequence = 0 " & _
         " ORDER BY tblTransaction.persID, tblTransactionSub.trnsDate, tblTransaction.trnPageCounter; "
End If

Dim Db As Database
Dim rs As DAO.Recordset
Set Db = CurrentDb()
Set rs = Db.OpenRecordset(strSQL)
If rs.RecordCount > 0 Then
    rs.MoveFirst
    Do Until rs.EOF
        rs.Edit
        rs.Fields("trnsSequence") = i
        rs.Update
        rs.MoveNext
        i = i + 1
    Loop
End If
rs.Close
Db.Close

fsSequence = True
End Function
Function FnIsLoad(ByVal FormName As String) As Boolean
Const conObjStateClosed = 0
Const conDesignView = 0
    If SysCmd(acSysCmdGetObjectState, acForm, FormName) <> conObjStateClosed Then
        If Forms(FormName).CurrentView <> conDesignView Then
        FnIsLoad = True
        End If
    End If
End Function