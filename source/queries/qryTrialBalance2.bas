dbMemo "SQL" ="SELECT qryChartWithAccounts.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartWi"
    "thAccounts.coaRef2,0) AS coaRef2, IIf(Forms.frmReport.txtToStr2>=3,qryChartWithA"
    "ccounts.coaRef3,0) AS coaRef3, IIf(Forms.frmReport.txtToStr2>=4,qryChartWithAcco"
    "unts.coaRef4,0) AS coaRef4, IIf(Forms.frmReport.txtToStr2>=5,qryChartWithAccount"
    "s.coaRef5,0) AS coaRef5, IIf(IsNumeric(Sum(qryTrialBalance1.Bal)),Sum(qryTrialBa"
    "lance1.Bal),0) AS Bal, IIf(Forms.frmReport.txtFromStr1=0,Sum(qryChartWithAccount"
    "s.accOpenBalance),0) AS OpenBal, Sum(qryTrialBalance1.CR) AS CR, Sum(qryTrialBal"
    "ance1.DR) AS DR, IIf(Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accID) AS "
    "accID, IIf(Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accNo) AS accNo, IIf"
    "(Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accName) AS accName, IIf(Forms"
    ".frmReport.txtToStr2=5,qryChartWithAccounts.accType) AS accType\015\012FROM qryT"
    "rialBalance1 RIGHT JOIN qryChartWithAccounts ON qryTrialBalance1.caoID=qryChartW"
    "ithAccounts.coaID\015\012WHERE qryChartWithAccounts.accType In (1,2,3)\015\012GR"
    "OUP BY qryChartWithAccounts.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartWi"
    "thAccounts.coaRef2,0), IIf(Forms.frmReport.txtToStr2>=3,qryChartWithAccounts.coa"
    "Ref3,0), IIf(Forms.frmReport.txtToStr2>=4,qryChartWithAccounts.coaRef4,0), IIf(F"
    "orms.frmReport.txtToStr2>=5,qryChartWithAccounts.coaRef5,0), IIf(Forms.frmReport"
    ".txtToStr2=5,qryChartWithAccounts.accID), IIf(Forms.frmReport.txtToStr2=5,qryCha"
    "rtWithAccounts.accNo), IIf(Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accN"
    "ame), IIf(Forms.frmReport.txtToStr2=5,qryChartWithAccounts.accType)\015\012ORDER"
    " BY qryChartWithAccounts.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChartWithA"
    "ccounts.coaRef2,0), IIf(Forms.frmReport.txtToStr2>=3,qryChartWithAccounts.coaRef"
    "3,0), IIf(Forms.frmReport.txtToStr2>=4,qryChartWithAccounts.coaRef4,0), IIf(Form"
    "s.frmReport.txtToStr2>=5,qryChartWithAccounts.coaRef5,0);\015\012"
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
