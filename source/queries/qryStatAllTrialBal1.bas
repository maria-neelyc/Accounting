dbMemo "SQL" ="SELECT Sum(tblTransactionSub.trnsCredits) AS CR, Sum(tblTransactionSub.trnsDebit"
    "s) AS DR, Sum(IIf(tblTransactionSub.trnsSign=\"C\",tblTransactionSub.trnsFAmount"
    ",0)) AS FCR, Sum(IIf(tblTransactionSub.trnsSign=\"D\",tblTransactionSub.trnsFAmo"
    "unt,0)) AS FDR, tblAccount.accID, tblAccount.caoID, tblAccount.accNo, tblAccount"
    ".accName, qryAccBalAll.Bal AS Bal, qryAccBalAll.FBal\015\012FROM (((tblTransacti"
    "on INNER JOIN (tblPeriodSel INNER JOIN tblPeriod ON tblPeriodSel.perID=tblPeriod"
    ".perID) ON tblTransaction.persID=tblPeriodSel.persID) INNER JOIN tblControl ON t"
    "blTransaction.trnYear=tblControl.yearID) INNER JOIN ((tblAccount INNER JOIN qryA"
    "ccBalForeign ON tblAccount.accID=qryAccBalForeign.accID) INNER JOIN tblTransacti"
    "onSub ON tblAccount.accID=tblTransactionSub.accID) ON tblTransaction.trnID=tblTr"
    "ansactionSub.trnID) INNER JOIN qryAccBalAll ON tblAccount.accID=qryAccBalAll.acc"
    "ID\015\012WHERE (((tblAccount.accStatus)=True) And ((tblTransaction.trnYear)=tbl"
    "Control.yearID) And ((tblPeriod.perID) Between Forms.frmReport.txtFromStr2 And F"
    "orms.frmReport.txtToStr2) And ((tblAccount.accNo) Between Forms.frmReport.cboFro"
    "mStr1 And Forms.frmReport.cboToStr1))\015\012GROUP BY tblAccount.accID, tblAccou"
    "nt.caoID, tblAccount.accNo, tblAccount.accName, qryAccBalAll.Bal, qryAccBalAll.F"
    "Bal, tblAccount.accOpenBalance;\015\012"
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
    Begin
        dbText "Name" ="FCR"
    End
    Begin
        dbText "Name" ="FDR"
    End
End
