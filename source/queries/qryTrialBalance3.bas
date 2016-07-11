dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & tblChar"
    "t.coaRef1,2) & \"-\" & Right(Chr(48) & tblChart.coaRef2,2) & \"-\" & Right(Chr(4"
    "8) & tblChart.coaRef3,2) & \"-\" & Right(Chr(48) & tblChart.coaRef4,2) & \"-\" &"
    " Right(Chr(48) & Chr(48) & tblChart.coaRef5,3) AS coaFRef, Choose(tblChart.lvlID"
    ",Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFRef,11),coaFRef) AS co"
    "aRef, tblChart.lvlID, qryTrialBalance2.accID, IIf((IsNumeric(qryTrialBalance2.CR"
    ")),qryTrialBalance2.Bal,qryTrialBalance2.OpenBal) AS Bal, qryTrialBalance2.CR, q"
    "ryTrialBalance2.DR, IIf(qryTrialBalance2.OpenBal>=0,qryTrialBalance2.OpenBal,0) "
    "AS OpenBalDR, IIf(qryTrialBalance2.OpenBal<0,qryTrialBalance2.OpenBal,0) AS Open"
    "BalCR, qryTrialBalance2.accNo, qryTrialBalance2.accName, qryTrialBalance2.accTyp"
    "e, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblCh"
    "art.coaRef5, tblChart.coaAccType\015\012FROM tblChart INNER JOIN qryTrialBalance"
    "2 ON (tblChart.coaRef1=qryTrialBalance2.coaRef1) AND (tblChart.coaRef2=qryTrialB"
    "alance2.coaRef2) AND (tblChart.coaRef3=qryTrialBalance2.coaRef3) AND (tblChart.c"
    "oaRef4=qryTrialBalance2.coaRef4) AND (tblChart.coaRef5=qryTrialBalance2.coaRef5)"
    "\015\012WHERE (((tblChart.lvlID)<=Forms.frmReport.txtToStr2)) AND\015\012Right(C"
    "hr(48) & tblChart.coaRef1,2)& Right(Chr(48) & tblChart.coaRef2,2) & Right(Chr(48"
    ") & tblChart.coaRef3,2)  & Right(Chr(48) & tblChart.coaRef4,2) & Right(Chr(48) &"
    " Chr(48) & tblChart.coaRef5,3) BETWEEN Forms.frmReport.FromRef AND Forms.frmRepo"
    "rt.ToRef\015\012ORDER BY tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, t"
    "blChart.coaRef4, tblChart.coaRef5;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
Begin
    Begin
        dbText "Name" ="coaRef"
    End
    Begin
        dbText "Name" ="coaFRef"
        dbInteger "ColumnWidth" ="3795"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OpenBalDR"
    End
    Begin
        dbText "Name" ="OpenBalCR"
    End
    Begin
        dbText "Name" ="Bal"
    End
End
