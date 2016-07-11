dbMemo "SQL" ="SELECT tblTransaction.trnEntryDate, tblControl.yearID, tblTransaction.persID, tb"
    "lAccount.accNo, tblAccount.accName, tblTransactionSub.trnsDocDate, tblTransactio"
    "nSub.trnsNote, Format([tblTransactionSub.trnsDebits],'Currency') AS trnsDebits, "
    "Format([tblTransactionSub.trnsCredits],'Currency') AS trnsCredits, tblTransactio"
    "n.trnInternalRef, tblReference.refName, tblDescription.desName, tblAccount.accOp"
    "enBalance, tblTransactionSub.trnsDate, tblTransactionSub.docNo\015\012FROM (tblT"
    "ransaction INNER JOIN tblControl ON tblTransaction.trnYear=tblControl.yearID) IN"
    "NER JOIN (tblAccount INNER JOIN ((tblTransactionSub LEFT JOIN tblReference ON tb"
    "lTransactionSub.refID=tblReference.refID) LEFT JOIN tblDescription ON tblTransac"
    "tionSub.desID=tblDescription.desID) ON tblAccount.accID=tblTransactionSub.accID)"
    " ON tblTransaction.trnID=tblTransactionSub.trnID\015\012WHERE (((tblTransaction."
    "trnEntryDate) Between Forms.frmReport.txtFromDate1 And Forms.frmReport.txtToDate"
    "1) And ((tblAccount.accNo) Between Forms.frmReport.txtFromStr1 And Forms.frmRepo"
    "rt.txtToStr1))\015\012ORDER BY tblAccount.accNo;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="trnsDebits"
    End
    Begin
        dbText "Name" ="trnsCredits"
    End
End
