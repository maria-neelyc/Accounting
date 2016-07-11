dbMemo "SQL" ="SELECT qryChartWithAccountsVAT.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryChar"
    "tWithAccountsVAT.coaRef2,0) AS coaRef2, IIf(Forms.frmReport.txtToStr2>=3,qryChar"
    "tWithAccountsVAT.coaRef3,0) AS coaRef3, IIf(Forms.frmReport.txtToStr2>=4,qryChar"
    "tWithAccountsVAT.coaRef4,0) AS coaRef4, IIf(Forms.frmReport.txtToStr2>=5,qryChar"
    "tWithAccountsVAT.coaRef5,0) AS coaRef5, IIf(Sum(qryTrialBalance1.Bal)<>0,Sum(qry"
    "TrialBalance1.Bal),qryChartWithAccountsVAT.accOpenBalance) AS Bal, Sum(qryTrialB"
    "alance1.CR) AS CR, Sum(qryTrialBalance1.DR) AS DR, IIf(Forms.frmReport.txtToStr2"
    "=5,qryChartWithAccountsVAT.accID) AS accID, IIf(Forms.frmReport.txtToStr2=5,qryC"
    "hartWithAccountsVAT.accNo) AS accNo, IIf(Forms.frmReport.txtToStr2=5,qryChartWit"
    "hAccountsVAT.accName) AS accName, qryChartWithAccountsVAT.VatType\015\012FROM qr"
    "yTrialBalance1 RIGHT JOIN qryChartWithAccountsVAT ON qryTrialBalance1.caoID=qryC"
    "hartWithAccountsVAT.coaID\015\012WHERE qryChartWithAccountsVAT.accIsVat=-1\015\012"
    "GROUP BY qryChartWithAccountsVAT.coaRef1, IIf(Forms.frmReport.txtToStr2>=2,qryCh"
    "artWithAccountsVAT.coaRef2,0), IIf(Forms.frmReport.txtToStr2>=3,qryChartWithAcco"
    "untsVAT.coaRef3,0), IIf(Forms.frmReport.txtToStr2>=4,qryChartWithAccountsVAT.coa"
    "Ref4,0), IIf(Forms.frmReport.txtToStr2>=5,qryChartWithAccountsVAT.coaRef5,0), II"
    "f(Forms.frmReport.txtToStr2=5,qryChartWithAccountsVAT.accID), IIf(Forms.frmRepor"
    "t.txtToStr2=5,qryChartWithAccountsVAT.accNo), IIf(Forms.frmReport.txtToStr2=5,qr"
    "yChartWithAccountsVAT.accName), qryChartWithAccountsVAT.accOpenBalance, qryChart"
    "WithAccountsVAT.VatType\015\012ORDER BY qryChartWithAccountsVAT.coaRef1, IIf(For"
    "ms.frmReport.txtToStr2>=2,qryChartWithAccountsVAT.coaRef2,0), IIf(Forms.frmRepor"
    "t.txtToStr2>=3,qryChartWithAccountsVAT.coaRef3,0), IIf(Forms.frmReport.txtToStr2"
    ">=4,qryChartWithAccountsVAT.coaRef4,0), IIf(Forms.frmReport.txtToStr2>=5,qryChar"
    "tWithAccountsVAT.coaRef5,0);\015\012"
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
End
