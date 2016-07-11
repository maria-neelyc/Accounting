dbMemo "SQL" ="SELECT Sum(tblTransactionSub.trnsCredits) AS CR, Sum(tblTransactionSub.trnsDebit"
    "s) AS DR, IIf(Forms.frmReport.txtFromStr1=0,Sum(tblTransactionSub.trnsDebits-tbl"
    "TransactionSub.trnsCredits)+tblAccount.accOpenBalance,Sum(tblTransactionSub.trns"
    "Debits-tblTransactionSub.trnsCredits)) AS Bal, tblAccount.accID, tblAccount.caoI"
    "D, tblAccount.accNo, tblAccount.accName\015\012FROM (tblAccount INNER JOIN (tblT"
    "ransactionSub INNER JOIN (tblTransaction INNER JOIN (tblPeriodSel INNER JOIN tbl"
    "Period ON tblPeriodSel.perID=tblPeriod.perID) ON tblTransaction.persID=tblPeriod"
    "Sel.persID) ON tblTransactionSub.trnID=tblTransaction.trnID) ON tblAccount.accID"
    "=tblTransactionSub.accID) INNER JOIN tblControl ON tblTransaction.trnYear=tblCon"
    "trol.yearID\015\012WHERE (((tblPeriod.perID) Between Forms.frmReport.txtFromStr1"
    " And Forms.frmReport.txtToStr1) And ((tblAccount.accStatus)=True) And ((tblTrans"
    "action.trnYear)=tblControl.yearID)) And tblAccount.accType=1\015\012GROUP BY tbl"
    "Account.accID, tblAccount.caoID, tblAccount.accNo, tblAccount.accName, tblAccoun"
    "t.accOpenBalance;\015\012"
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
        dbText "Name" ="CR"
    End
    Begin
        dbText "Name" ="DR"
    End
End
