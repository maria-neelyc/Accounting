dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & tblChar"
    "t.coaRef1,2) & \"-\" & Right(Chr(48) & tblChart.coaRef2,2) & \"-\" & Right(Chr(4"
    "8) & tblChart.coaRef3,2) & \"-\" & Right(Chr(48) & tblChart.coaRef4,2) & \"-\" &"
    " Right(Chr(48) & Chr(48) & tblChart.coaRef5,3) AS coaFRef, Choose(tblChart.lvlID"
    ",Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFRef,11),coaFRef) AS co"
    "aRef, tblChart.lvlID, qryVatBal1.accID, qryVatBal1.Bal, qryVatBal1.CR, qryVatBal"
    "1.DR, qryVatBal1.accNo, qryVatBal1.accName, qryVatBal1.VatType\015\012FROM tblCh"
    "art INNER JOIN qryVatBal1 ON (qryVatBal1.coaRef5=tblChart.coaRef5) AND (qryVatBa"
    "l1.coaRef4=tblChart.coaRef4) AND (qryVatBal1.coaRef3=tblChart.coaRef3) AND (qryV"
    "atBal1.coaRef2=tblChart.coaRef2) AND (tblChart.coaRef1=qryVatBal1.coaRef1)\015\012"
    "WHERE (((tblChart.lvlID)<=Forms.frmReport.txtToStr2))\015\012ORDER BY qryVatBal1"
    ".VatType, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4"
    ", tblChart.coaRef5;\015\012"
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
    Begin
        dbText "Name" ="coaFRef"
        dbInteger "ColumnWidth" ="3795"
        dbBoolean "ColumnHidden" ="0"
    End
End
