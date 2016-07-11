dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaName, tblChart.coaNo, tblChart.coaIDg, tblCha"
    "rt.lvlID, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4"
    ", tblChart.coaRef5\015\012FROM tblChart\015\012WHERE tblChart.lvlID = 5\015\012O"
    "RDER BY tblChart.coaName, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, "
    "tblChart.coaRef4, tblChart.coaRef5;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
