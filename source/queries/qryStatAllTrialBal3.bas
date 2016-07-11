dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & tblChar"
    "t.coaRef1,2) & \"-\" & Right(Chr(48) & tblChart.coaRef2,2) & \"-\" & Right(Chr(4"
    "8) & tblChart.coaRef3,2) & \"-\" & Right(Chr(48) & tblChart.coaRef4,2) & \"-\" &"
    " Right(Chr(48) & Chr(48) & tblChart.coaRef5,3) AS coaFRef, Choose(tblChart.lvlID"
    ",Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFRef,11),coaFRef) AS co"
    "aRef, tblChart.lvlID, qryStatAllTrialBal2.accID, IIf((IsNumeric(qryStatAllTrialB"
    "al2.CR)),qryStatAllTrialBal2.Bal,qryStatAllTrialBal2.OpenBal) AS Bal, qryStatAll"
    "TrialBal2.CR, qryStatAllTrialBal2.DR, ABS(IIf(qryStatAllTrialBal2.OpenBal>=0,qry"
    "StatAllTrialBal2.OpenBal,0)) AS OpenBalDR, ABS(IIf(qryStatAllTrialBal2.OpenBal<0"
    ",qryStatAllTrialBal2.OpenBal,0)) AS OpenBalCR, IIf((IsNumeric(qryStatAllTrialBal"
    "2.FCR)),qryStatAllTrialBal2.FBal,qryStatAllTrialBal2.FOpenBal) AS FBal, qryStatA"
    "llTrialBal2.FCR, qryStatAllTrialBal2.FDR, IIf(qryStatAllTrialBal2.FOpenBal>=0,qr"
    "yStatAllTrialBal2.FOpenBal,0) AS FOpenBalDR, IIf(qryStatAllTrialBal2.FOpenBal<0,"
    "qryStatAllTrialBal2.FOpenBal,0) AS FOpenBalCR, qryStatAllTrialBal2.accNo, qrySta"
    "tAllTrialBal2.accName, qryStatAllTrialBal2.accCur\015\012FROM tblChart INNER JOI"
    "N qryStatAllTrialBal2 ON (tblChart.coaRef5=qryStatAllTrialBal2.coaRef5) AND (tbl"
    "Chart.coaRef4=qryStatAllTrialBal2.coaRef4) AND (tblChart.coaRef3=qryStatAllTrial"
    "Bal2.coaRef3) AND (tblChart.coaRef2=qryStatAllTrialBal2.coaRef2) AND (tblChart.c"
    "oaRef1=qryStatAllTrialBal2.coaRef1)\015\012WHERE (((tblChart.lvlID)=5))\015\012O"
    "RDER BY tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, "
    "tblChart.coaRef5;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
Begin
    Begin
        dbText "Name" ="coaFRef"
    End
    Begin
        dbText "Name" ="coaRef"
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
    Begin
        dbText "Name" ="FBal"
    End
    Begin
        dbText "Name" ="FOpenBalDR"
    End
    Begin
        dbText "Name" ="FOpenBalCR"
    End
End
