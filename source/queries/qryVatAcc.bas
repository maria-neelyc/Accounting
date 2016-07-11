dbMemo "SQL" ="SELECT DISTINCT tblAccount.accNo, \"INPUT\" As VatType\015\012FROM tblAccount IN"
    "NER JOIN tblVat ON tblAccount.accID = tblVat.vatInput\015\012WHERE (((tblVat.vat"
    "Input)>0))\015\012UNION SELECT DISTINCT tblAccount.accNo,  \"OUTPUT\" As VatType"
    "\015\012FROM tblAccount INNER JOIN tblVat ON tblAccount.accID = tblVat.vatOutput"
    "\015\012WHERE (((tblVat.vatOutput)>0))\015\012UNION SELECT DISTINCT tblAccount.a"
    "ccNo, \"CONTROLT\" AS VatType\015\012FROM tblControl INNER JOIN tblAccount ON tb"
    "lControl.ctrControlAcc = tblAccount.accID;\015\012"
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
