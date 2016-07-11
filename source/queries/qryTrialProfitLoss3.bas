dbMemo "SQL" ="SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & tblChar"
    "t.coaRef1,2) & \"-\" & Right(Chr(48) & tblChart.coaRef2,2) & \"-\" & Right(Chr(4"
    "8) & tblChart.coaRef3,2) & \"-\" & Right(Chr(48) & tblChart.coaRef4,2) & \"-\" &"
    " Right(Chr(48) & Chr(48) & tblChart.coaRef5,3) AS coaFRef, Choose(tblChart.lvlID"
    ",Left(coaFRef,2),Left(coaFRef,5),Left(coaFRef,8),Left(coaFRef,11),coaFRef) AS co"
    "aRef, tblChart.lvlID, qryTrialProfitLoss2.accID, IIf((IsNumeric(qryTrialProfitLo"
    "ss2.CR)),qryTrialProfitLoss2.Bal,qryTrialProfitLoss2.OpenBal) AS Bal, qryTrialPr"
    "ofitLoss2.CR, qryTrialProfitLoss2.DR, IIf(qryTrialProfitLoss2.OpenBal>=0,qryTria"
    "lProfitLoss2.OpenBal,0) AS OpenBalDR, IIf(qryTrialProfitLoss2.OpenBal<0,qryTrial"
    "ProfitLoss2.OpenBal,0) AS OpenBalCR, qryTrialProfitLoss2.accNo, qryTrialProfitLo"
    "ss2.accName, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaR"
    "ef4, tblChart.coaRef5\015\012FROM tblChart INNER JOIN qryTrialProfitLoss2 ON (qr"
    "yTrialProfitLoss2.coaRef5=tblChart.coaRef5) AND (qryTrialProfitLoss2.coaRef4=tbl"
    "Chart.coaRef4) AND (qryTrialProfitLoss2.coaRef3=tblChart.coaRef3) AND (qryTrialP"
    "rofitLoss2.coaRef2=tblChart.coaRef2) AND (tblChart.coaRef1=qryTrialProfitLoss2.c"
    "oaRef1)\015\012WHERE (((tblChart.lvlID)<=Forms.frmReport.txtToStr2)) AND\015\012"
    " Right(Chr(48) & tblChart.coaRef1,2) &Right(Chr(48) & tblChart.coaRef2,2) &Right"
    "(Chr(48) & tblChart.coaRef3,2) &Right(Chr(48) & tblChart.coaRef4,2) & Right(Chr("
    "48) & Chr(48) & tblChart.coaRef5,3) BETWEEN Forms.frmReport.FromRef AND Forms.fr"
    "mReport.ToRef\015\012ORDER BY tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRe"
    "f3, tblChart.coaRef4, tblChart.coaRef5;\015\012"
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
