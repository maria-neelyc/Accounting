dbMemo "SQL" ="SELECT Sum(tblTransactionSub.trnsCredits) AS CR, Sum(tblTransactionSub.trnsDebit"
    "s) AS DR, tblAccount.accID, tblAccount.caoID, tblAccount.accNo, tblAccount.accNa"
    "me, qryAccountBalances.Bal\015\012FROM ((tblTransaction INNER JOIN (tblPeriodSel"
    " INNER JOIN tblPeriod ON tblPeriodSel.perID=tblPeriod.perID) ON tblTransaction.p"
    "ersID=tblPeriodSel.persID) INNER JOIN tblControl ON tblTransaction.trnYear=tblCo"
    "ntrol.yearID) INNER JOIN ((tblAccount INNER JOIN qryAccountBalances ON tblAccoun"
    "t.accID=qryAccountBalances.accID) INNER JOIN tblTransactionSub ON tblAccount.acc"
    "ID=tblTransactionSub.accID) ON tblTransaction.trnID=tblTransactionSub.trnID\015\012"
    "WHERE (((tblAccount.accStatus)=True) And ((tblTransaction.trnYear)=tblControl.ye"
    "arID) And ((tblPeriod.perID) Between Forms.frmReport.txtFromStr2 And Forms.frmRe"
    "port.txtToStr2) And ((tblAccount.accNo) Between Forms.frmReport.cboFromStr1 And "
    "Forms.frmReport.cboToStr1))\015\012GROUP BY tblAccount.accID, tblAccount.caoID, "
    "tblAccount.accNo, tblAccount.accName, qryAccountBalances.Bal, tblAccount.accOpen"
    "Balance;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="CR"
    End
    Begin
        dbText "Name" ="DR"
    End
End
