dbMemo "SQL" ="SELECT qryChartWithAccounts.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartWi"
    "thAccounts.coaRef2,0) AS coaRef2, IIf(Forms.frmReport.txtToStr2>=3,qryChartWithA"
    "ccounts.coaRef3,0) AS coaRef3, IIf(Forms.frmReport.txtToStr2>=4,qryChartWithAcco"
    "unts.coaRef4,0) AS coaRef4, IIf(Forms.frmReport.txtToStr2>=5,qryChartWithAccount"
    "s.coaRef5,0) AS coaRef5, Sum(qryTrialProfitLoss1.Bal) AS Bal, IIf(Forms.frmRepor"
    "t.txtFromStr1=0,Sum(qryChartWithAccounts.accOpenBalance),0) AS OpenBal, Sum(qryT"
    "rialProfitLoss1.CR) AS CR, Sum(qryTrialProfitLoss1.DR) AS DR, IIf(Forms.frmRepor"
    "t.txtToStr2=5,qryChartWithAccounts.accID) AS accID, IIf(Forms.frmReport.txtToStr"
    "2=5,qryChartWithAccounts.accNo) AS accNo, IIf(Forms.frmReport.txtToStr2=5,qryCha"
    "rtWithAccounts.accName) AS accName\015\012FROM qryTrialProfitLoss1 RIGHT JOIN qr"
    "yChartWithAccounts ON qryTrialProfitLoss1.caoID=qryChartWithAccounts.coaID\015\012"
    "WHERE qryChartWithAccounts.accType In (1,3)\015\012GROUP BY qryChartWithAccounts"
    ".coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartWithAccounts.coaRef2,0), IIf("
    "Forms.frmReport.txtToStr2>=3,qryChartWithAccounts.coaRef3,0), IIf(Forms.frmRepor"
    "t.txtToStr2>=4,qryChartWithAccounts.coaRef4,0), IIf(Forms.frmReport.txtToStr2>=5"
    ",qryChartWithAccounts.coaRef5,0), IIf(Forms.frmReport.txtToStr2=5,qryChartWithAc"
    "counts.accID), IIf(Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accNo), IIf("
    "Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accName)\015\012ORDER BY qryCha"
    "rtWithAccounts.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartWithAccounts.co"
    "aRef2,0), IIf(Forms.frmReport.txtToStr2>=3,qryChartWithAccounts.coaRef3,0), IIf("
    "Forms.frmReport.txtToStr2>=4,qryChartWithAccounts.coaRef4,0), IIf(Forms.frmRepor"
    "t.txtToStr2>=5,qryChartWithAccounts.coaRef5,0);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="accID"
    End
    Begin
        dbText "Name" ="accNo"
    End
    Begin
        dbText "Name" ="accName"
    End
    Begin
        dbText "Name" ="coaRef2"
    End
    Begin
        dbText "Name" ="coaRef3"
    End
    Begin
        dbText "Name" ="coaRef4"
    End
    Begin
        dbText "Name" ="coaRef5"
    End
    Begin
        dbText "Name" ="CR"
    End
    Begin
        dbText "Name" ="DR"
    End
    Begin
        dbText "Name" ="Bal"
    End
    Begin
        dbText "Name" ="OpenBal"
    End
End
