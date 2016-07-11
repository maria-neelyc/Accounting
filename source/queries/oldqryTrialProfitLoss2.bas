dbMemo "SQL" ="SELECT qryLeveledChart.coaRef*(Choose(Forms.frmReport.txtToStr2,1000000000,10000"
    "000,100000,1000)) AS coaRef, Sum(qryTrialProfitLoss1.Bal) AS Bal, IIf(Forms.frmR"
    "eport.txtToStr2=4,qryTrialProfitLoss1.accID) AS accID, IIf(Forms.frmReport.txtTo"
    "Str2=4,qryTrialProfitLoss1.accNo) AS accNo, IIf(Forms.frmReport.txtToStr2=4,qryT"
    "rialProfitLoss1.accName) AS accName\015\012FROM qryTrialProfitLoss1 RIGHT JOIN q"
    "ryLeveledChart ON qryTrialProfitLoss1.caoID=qryLeveledChart.coaID\015\012GROUP B"
    "Y IIf(Forms.frmReport.txtToStr2=4,qryTrialProfitLoss1.accID), qryLeveledChart.co"
    "aRef, IIf(Forms.frmReport.txtToStr2=4,qryTrialProfitLoss1.accNo), IIf(Forms.frmR"
    "eport.txtToStr2=4,qryTrialProfitLoss1.accName)\015\012ORDER BY qryLeveledChart.c"
    "oaRef;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="coaRef"
    End
    Begin
        dbText "Name" ="Bal"
    End
    Begin
        dbText "Name" ="accID"
    End
    Begin
        dbText "Name" ="accNo"
    End
    Begin
        dbText "Name" ="accName"
    End
End
