dbMemo "SQL" ="SELECT Sum(tblTransactionSub.trnsDebits-tblTransactionSub.trnsCredits)+tblAccoun"
    "t.accOpenBalance AS Bal, Sum(IIf(tblTransactionSub.trnsSign=\"D\",tblTransaction"
    "Sub.trnsFAmount,-tblTransactionSub.trnsFAmount))+tblAccount.accOpenBalCur AS FBa"
    "l, tblAccount.accID\015\012FROM ((tblTransaction INNER JOIN (tblPeriodSel INNER "
    "JOIN tblPeriod ON tblPeriodSel.perID=tblPeriod.perID) ON tblTransaction.persID=t"
    "blPeriodSel.persID) INNER JOIN tblControl ON tblTransaction.trnYear=tblControl.y"
    "earID) INNER JOIN (tblAccount INNER JOIN tblTransactionSub ON tblAccount.accID=t"
    "blTransactionSub.accID) ON tblTransaction.trnID=tblTransactionSub.trnID\015\012W"
    "HERE (((tblAccount.accStatus)=True) And ((tblTransaction.trnYear)=tblControl.yea"
    "rID))\015\012GROUP BY tblAccount.accID, tblAccount.accOpenBalance, tblAccount.ac"
    "cOpenBalCur;\015\012"
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
    Begin
        dbText "Name" ="FBal"
    End
End
