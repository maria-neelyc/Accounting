dbMemo "SQL" ="SELECT tblTransaction.trnEntryDate, tblControl.yearID, tblTransaction.persID, qr"
    "yChartWithAccountsVAT.accNo, qryChartWithAccountsVAT.accName, tblTransactionSub."
    "trnsNote, tblTransactionSub.trnsDebits, tblTransactionSub.trnsCredits, tblTransa"
    "ction.trnInternalRef, tblReference.refName, tblDescription.desName, qryChartWith"
    "AccountsVAT.accOpenBalance, tblTransactionSub.trnsDate, qryChartWithAccountsVAT."
    "VatType\015\012FROM ((((tblTransaction INNER JOIN tblTransactionSub ON tblTransa"
    "ction.trnID=tblTransactionSub.trnID) LEFT JOIN tblReference ON tblTransactionSub"
    ".refID=tblReference.refID) LEFT JOIN tblDescription ON tblTransactionSub.desID=t"
    "blDescription.desID) INNER JOIN tblControl ON tblTransaction.trnYear=tblControl."
    "yearID) INNER JOIN qryChartWithAccountsVAT ON tblTransactionSub.accID=qryChartWi"
    "thAccountsVAT.accID\015\012WHERE (((tblTransaction.trnEntryDate) Between Format("
    "Forms.frmReport.txtFromDate1,\"yyyy-dd-mm\") And Format(Forms.frmReport.txtToDat"
    "e1,\"yyyy-dd-mm\")) And ((qryChartWithAccountsVAT.accNo) Between Forms.frmReport"
    ".cboFromStr1 And Forms.frmReport.cboToStr1))\015\012ORDER BY qryChartWithAccount"
    "sVAT.VatType, qryChartWithAccountsVAT.accNo;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
