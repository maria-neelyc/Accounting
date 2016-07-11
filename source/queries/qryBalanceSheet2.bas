dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & tblChar"
    "t.coaRef1,2) & \"-\" & Right(Chr(48) & tblChart.coaRef2,2) & \"-\" & Right(Chr(4"
    "8) & tblChart.coaRef3,2) & \"-\" & Right(Chr(48) & tblChart.coaRef4,2) & \"-\" &"
    " Right(Chr(48) & Chr(48) & tblChart.coaRef5,3) AS coaFRef, Choose(tblChart.lvlID"
    ",Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFRef,11),coaFRef) AS co"
    "aRef, tblChart.lvlID, qryBalanceSheet1.accID, IIf((IsNumeric(qryBalanceSheet1.CR"
    ")),qryBalanceSheet1.Bal,qryBalanceSheet1.OpenBal) AS Bal, qryBalanceSheet1.CR, q"
    "ryBalanceSheet1.DR, IIf(qryBalanceSheet1.OpenBal>=0,qryBalanceSheet1.OpenBal,0) "
    "AS OpenBalDR, IIf(qryBalanceSheet1.OpenBal<0,qryBalanceSheet1.OpenBal,0) AS Open"
    "BalCR, qryBalanceSheet1.accNo, qryBalanceSheet1.accName, qryBalanceSheet1.accTyp"
    "e\015\012FROM tblChart INNER JOIN qryBalanceSheet1 ON (tblChart.coaRef1=qryBalan"
    "ceSheet1.coaRef1) AND (tblChart.coaRef2=qryBalanceSheet1.coaRef2) AND (tblChart."
    "coaRef3=qryBalanceSheet1.coaRef3) AND (tblChart.coaRef4=qryBalanceSheet1.coaRef4"
    ") AND (tblChart.coaRef5=qryBalanceSheet1.coaRef5)\015\012WHERE (((tblChart.lvlID"
    ")<=Forms.frmReport.txtToStr2))\015\012AND\015\012 Right(Chr(48) & tblChart.coaRe"
    "f1,2) &Right(Chr(48) & tblChart.coaRef2,2) &Right(Chr(48) & tblChart.coaRef3,2) "
    "&Right(Chr(48) & tblChart.coaRef4,2) & Right(Chr(48) & Chr(48) & tblChart.coaRef"
    "5,3) BETWEEN Forms.frmReport.FromRef AND Forms.frmReport.ToRef\015\012ORDER BY t"
    "blChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart."
    "coaRef5;\015\012"
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
End
