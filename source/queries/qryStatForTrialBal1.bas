dbMemo "SQL" ="SELECT Sum(IIf(tblTransactionSub.trnsSign=\"C\",tblTransactionSub.trnsFAmount,0)"
    ") AS CR, Sum(IIf(tblTransactionSub.trnsSign=\"D\",tblTransactionSub.trnsFAmount,"
    "0)) AS DR, tblAccount.accID, tblAccount.caoID, tblAccount.accNo, tblAccount.accN"
    "ame, qryAccBalForeign.Bal AS Bal\015\012FROM (((tblTransaction INNER JOIN (tblPe"
    "riodSel INNER JOIN tblPeriod ON tblPeriodSel.perID=tblPeriod.perID) ON tblTransa"
    "ction.persID=tblPeriodSel.persID) INNER JOIN tblControl ON tblTransaction.trnYea"
    "r=tblControl.yearID) INNER JOIN (tblAccount INNER JOIN tblTransactionSub ON tblA"
    "ccount.accID=tblTransactionSub.accID) ON tblTransaction.trnID=tblTransactionSub."
    "trnID) INNER JOIN qryAccBalForeign ON tblAccount.accID=qryAccBalForeign.accID\015"
    "\012WHERE (((tblAccount.accStatus)=True) And ((tblTransaction.trnYear)=tblContro"
    "l.yearID) And ((tblPeriod.perID) Between Forms.frmReport.txtFromStr2 And Forms.f"
    "rmReport.txtToStr2) And ((tblAccount.accNo) Between Forms.frmReport.cboFromStr1 "
    "And Forms.frmReport.cboToStr1))\015\012GROUP BY tblAccount.accID, tblAccount.cao"
    "ID, tblAccount.accNo, tblAccount.accName, qryAccBalForeign.Bal, tblAccount.accOp"
    "enBalance;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
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
