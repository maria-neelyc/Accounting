dbMemo "SQL" ="SELECT qryChartwithAccounts.coaID, qryChartwithAccounts.lvlID, qryChartwithAccou"
    "nts.coaIDg, qryChartwithAccounts.coaNo, qryChartwithAccounts.coaName, qryChartwi"
    "thAccounts.coaRef1, qryChartwithAccounts.coaRef2, qryChartwithAccounts.coaRef3, "
    "qryChartwithAccounts.coaRef4, qryChartwithAccounts.coaRef5, qryChartwithAccounts"
    ".accNo, qryChartwithAccounts.accName, qryChartwithAccounts.accID, qryChartwithAc"
    "counts.accIsVat, qryChartwithAccounts.accType, qryChartwithAccounts.accOpenBalan"
    "ce, 'INPUT' As VatType\015\012FROM qryChartwithAccounts INNER JOIN tblVat ON qry"
    "ChartwithAccounts.accID = tblVat.vatInput\015\012UNION SELECT qryChartwithAccoun"
    "ts.coaID, qryChartwithAccounts.lvlID, qryChartwithAccounts.coaIDg, qryChartwithA"
    "ccounts.coaNo, qryChartwithAccounts.coaName, qryChartwithAccounts.coaRef1, qryCh"
    "artwithAccounts.coaRef2, qryChartwithAccounts.coaRef3, qryChartwithAccounts.coaR"
    "ef4, qryChartwithAccounts.coaRef5, qryChartwithAccounts.accNo, qryChartwithAccou"
    "nts.accName, qryChartwithAccounts.accID, qryChartwithAccounts.accIsVat, qryChart"
    "withAccounts.accType, qryChartwithAccounts.accOpenBalance, 'OUTPUT' AS VatType\015"
    "\012FROM qryChartwithAccounts INNER JOIN tblVat ON qryChartwithAccounts.accID = "
    "tblVat.vatOutput\015\012UNION SELECT qryChartwithAccounts.coaID, qryChartwithAcc"
    "ounts.lvlID, qryChartwithAccounts.coaIDg, qryChartwithAccounts.coaNo, qryChartwi"
    "thAccounts.coaName, qryChartwithAccounts.coaRef1, qryChartwithAccounts.coaRef2, "
    "qryChartwithAccounts.coaRef3, qryChartwithAccounts.coaRef4, qryChartwithAccounts"
    ".coaRef5, qryChartwithAccounts.accNo, qryChartwithAccounts.accName, qryChartwith"
    "Accounts.accID, qryChartwithAccounts.accIsVat, qryChartwithAccounts.accType, qry"
    "ChartwithAccounts.accOpenBalance, 'CONTROL' AS VatType\015\012FROM tblControl IN"
    "NER JOIN qryChartwithAccounts ON tblControl.ctrControlAcc = qryChartwithAccounts"
    ".accID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="VatType"
    End
End
