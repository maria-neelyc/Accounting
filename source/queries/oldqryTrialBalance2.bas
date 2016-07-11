dbMemo "SQL" ="SELECT qryLeveledChart.coaRef*(Choose(Forms.frmReport.txtToStr2,1000000000,10000"
    "000,100000,1000,1)) AS coaRef, Sum(qryTrialBalance1.Bal) AS Bal, IIf(Forms.frmRe"
    "port.txtToStr2=5,qryTrialBalance1.accID) AS accID, IIf(Forms.frmReport.txtToStr2"
    "=5,qryTrialBalance1.accNo) AS accNo, IIf(Forms.frmReport.txtToStr2=5,qryTrialBal"
    "ance1.accName) AS accName\015\012FROM qryTrialBalance1 RIGHT JOIN qryLeveledChar"
    "t ON qryTrialBalance1.caoID=qryLeveledChart.coaID\015\012GROUP BY IIf(Forms.frmR"
    "eport.txtToStr2=5,qryTrialBalance1.accID), qryLeveledChart.coaRef, IIf(Forms.frm"
    "Report.txtToStr2=5,qryTrialBalance1.accNo), IIf(Forms.frmReport.txtToStr2=5,qryT"
    "rialBalance1.accName)\015\012ORDER BY qryLeveledChart.coaRef;\015\012"
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
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
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
