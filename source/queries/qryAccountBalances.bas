dbMemo "SQL" ="SELECT Sum(tblTransactionSub.trnsDebits-tblTransactionSub.trnsCredits)+tblAccoun"
    "t.accOpenBalance AS Bal, tblAccount.accID\015\012FROM (tblAccount INNER JOIN (tb"
    "lTransactionSub INNER JOIN (tblTransaction INNER JOIN (tblPeriodSel INNER JOIN t"
    "blPeriod ON tblPeriodSel.perID=tblPeriod.perID) ON tblTransaction.persID=tblPeri"
    "odSel.persID) ON tblTransactionSub.trnID=tblTransaction.trnID) ON tblAccount.acc"
    "ID=tblTransactionSub.accID) INNER JOIN tblControl ON tblTransaction.trnYear=tblC"
    "ontrol.yearID\015\012WHERE (((tblAccount.accStatus)=True) And ((tblTransaction.t"
    "rnYear)=tblControl.yearID) And ((tblPeriod.perID)<=Forms.frmReport.txtToStr2))\015"
    "\012GROUP BY tblAccount.accID, tblAccount.accOpenBalance;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Bal"
    End
End
