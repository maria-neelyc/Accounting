dbMemo "SQL" ="SELECT Sum(IIf(tblTransactionSub.trnsSign=\"D\",tblTransactionSub.trnsFAmount,-t"
    "blTransactionSub.trnsFAmount))+tblAccount.accOpenBalCur AS Bal, tblAccount.accID"
    "\015\012FROM ((tblTransaction INNER JOIN (tblPeriodSel INNER JOIN tblPeriod ON t"
    "blPeriodSel.perID=tblPeriod.perID) ON tblTransaction.persID=tblPeriodSel.persID)"
    " INNER JOIN tblControl ON tblTransaction.trnYear=tblControl.yearID) INNER JOIN ("
    "tblAccount INNER JOIN tblTransactionSub ON tblAccount.accID=tblTransactionSub.ac"
    "cID) ON tblTransaction.trnID=tblTransactionSub.trnID\015\012WHERE (((tblAccount."
    "accStatus)=True) And ((tblTransaction.trnYear)=tblControl.yearID))\015\012GROUP "
    "BY tblAccount.accID, tblAccount.accOpenBalCur;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
Begin
    Begin
        dbText "Name" ="Bal"
    End
End
