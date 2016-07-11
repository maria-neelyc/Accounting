dbMemo "SQL" ="SELECT qryChartwithAccounts.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartwi"
    "thAccounts.coaRef2,0) AS coaRef2, IIf(Forms.frmReport.txtToStr2>=3,qryChartwithA"
    "ccounts.coaRef3,0) AS coaRef3, IIf(Forms.frmReport.txtToStr2>=4,qryChartwithAcco"
    "unts.coaRef4,0) AS coaRef4, IIf(Forms.frmReport.txtToStr2>=5,qryChartwithAccount"
    "s.coaRef5,0) AS coaRef5, Sum(qryTrialBalance1.Bal) AS Bal, IIf(Forms.frmReport.t"
    "xtFromStr1=0,Sum(qryChartwithAccounts.accOpenBalance),0) AS OpenBal, Sum(qryTria"
    "lBalance1.CR) AS CR, Sum(qryTrialBalance1.DR) AS DR, IIf(Forms.frmReport.txtToSt"
    "r2=5,qryChartwithAccounts.accID) AS accID, IIf(Forms.frmReport.txtToStr2=5,qryCh"
    "artwithAccounts.accNo) AS accNo, IIf(Forms.frmReport.txtToStr2=5,qryChartwithAcc"
    "ounts.accName) AS accName, IIf(Forms.frmReport.txtToStr2=5,qryChartwithAccounts."
    "accType) AS accType\015\012FROM qryTrialBalance1 RIGHT JOIN qryChartwithAccounts"
    " ON qryTrialBalance1.caoID=qryChartwithAccounts.coaID\015\012WHERE qryChartWithA"
    "ccounts.accType In (2,3)\015\012GROUP BY qryChartwithAccounts.coaRef1, IIf(Forms"
    ".frmReport.txtToStr2>=2,qryChartwithAccounts.coaRef2,0), IIf(Forms.frmReport.txt"
    "ToStr2>=3,qryChartwithAccounts.coaRef3,0), IIf(Forms.frmReport.txtToStr2>=4,qryC"
    "hartwithAccounts.coaRef4,0), IIf(Forms.frmReport.txtToStr2>=5,qryChartwithAccoun"
    "ts.coaRef5,0), IIf(Forms.frmReport.txtToStr2=5,qryChartwithAccounts.accID), IIf("
    "Forms.frmReport.txtToStr2=5,qryChartwithAccounts.accNo), IIf(Forms.frmReport.txt"
    "ToStr2=5,qryChartwithAccounts.accName), IIf(Forms.frmReport.txtToStr2=5,qryChart"
    "withAccounts.accType)\015\012ORDER BY qryChartwithAccounts.coaRef1, IIf(Forms.fr"
    "mReport.txtToStr2>=2,qryChartwithAccounts.coaRef2,0), IIf(Forms.frmReport.txtToS"
    "tr2>=3,qryChartwithAccounts.coaRef3,0), IIf(Forms.frmReport.txtToStr2>=4,qryChar"
    "twithAccounts.coaRef4,0), IIf(Forms.frmReport.txtToStr2>=5,qryChartwithAccounts."
    "coaRef5,0);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
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
        dbText "Name" ="OpenBal"
    End
    Begin
        dbText "Name" ="accType"
    End
End
