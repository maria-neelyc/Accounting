dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(\"0\" & CStr(tblC"
    "hart.coaRef\\1000000000),2) & \"-\" & Right(\"0\" & CStr((tblChart.coaRef\\10000"
    "000) Mod 1000),2) & \"-\" & Right(\"0\" & CStr((tblChart.coaRef\\100000) Mod 100"
    "0),2) & \"-\" & Right(\"0\" & CStr((tblChart.coaRef\\1000) Mod 1000),2) & \"-\" "
    "& Right(\"00\" & CStr((tblChart.coaRef) Mod 1000),3) AS coaFRef, Choose(tblChart"
    ".lvlID,Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFRef,11),coaFRef)"
    " AS coaRef, tblChart.lvlID, qryTrialBalance2.accID, qryTrialBalance2.Bal, qryTri"
    "alBalance2.accNo, qryTrialBalance2.accName\015\012FROM tblChart INNER JOIN qryTr"
    "ialBalance2 ON tblChart.coaRef=qryTrialBalance2.coaRef\015\012WHERE tblChart.lvl"
    "ID=Forms.frmReport.txtToStr2\015\012ORDER BY tblChart.coaRef;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="coaRef"
    End
End
