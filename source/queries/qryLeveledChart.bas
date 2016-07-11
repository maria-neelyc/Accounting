dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.lvlID, tblChart.coaIDg, tblChart.coaNo, tblChart"
    ".coaName, Choose(Forms.frmReport.txtToStr2,Right(Chr(48) & tblChart.coaRef1,2),R"
    "ight(Chr(48) & tblChart.coaRef1,2) & Right(Chr(48) & tblChart.coaRef2,2),Right(C"
    "hr(48) & tblChart.coaRef1,2) & Right(Chr(48) & tblChart.coaRef2,2) & Right(Chr(4"
    "8) & tblChart.coaRef3,2),Right(Chr(48) & tblChart.coaRef1,2) & Right(Chr(48) & t"
    "blChart.coaRef2,2) & Right(Chr(48) & tblChart.coaRef3,2) & Right(Chr(48) & tblCh"
    "art.coaRef4,2),Right(Chr(48) & tblChart.coaRef1,2) & Right(Chr(48) & tblChart.co"
    "aRef2,2) & Right(Chr(48) & tblChart.coaRef3,2) & Right(Chr(48) & tblChart.coaRef"
    "4,2) & Right(Chr(48) & Chr(48) & tblChart.coaRef5,3)) AS coaRef\015\012FROM tblC"
    "hart;\015\012"
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
