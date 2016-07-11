dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblAccount.accNo, tblChart.coaName AS coa"
    "Name, Right(Chr(48) & tblChart.coaRef1,2) & \" - \" & Right(Chr(48) & tblChart.c"
    "oaRef2,2) & \" - \" & Right(Chr(48) & tblChart.coaRef3,2) & \" - \" & Right(Chr("
    "48) & tblChart.coaRef4,2) & \" - \" & Right(Chr(48) & Chr(48) & tblChart.coaRef5"
    ",3) AS coaFRef, Choose(tblChart.lvlID,Left(coaFRef,2),Left(coaFRef,7),Left(coaFR"
    "ef,12),Left(coaFRef,17),coaFRef) AS coaRef, tblChart.lvlID\015\012FROM tblChart "
    "LEFT JOIN tblAccount ON tblChart.coaID=tblAccount.caoID\015\012ORDER BY tblChart"
    ".coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart.coaRef5"
    ";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="tblChart.coaNo"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="coaRef"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="coaFRef"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="coaName"
        dbInteger "ColumnWidth" ="5235"
        dbBoolean "ColumnHidden" ="0"
    End
End
