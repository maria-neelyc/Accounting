dbMemo "SQL" ="SELECT Sum(IIf([trnsSign]='C',[trnsAmount],-1*[trnsAmount]))+prevtblAccount.accO"
    "penBalance AS Bal, prevtblAccount.accNo, prevtblAccount.caoID, prevtblAccount.ac"
    "cName\015\012FROM prevtblAccount INNER JOIN (prevtblTransactionSub INNER JOIN pr"
    "evtblTransaction ON prevtblTransactionSub.trnID = prevtblTransaction.trnID) ON p"
    "revtblAccount.accID = prevtblTransactionSub.accID\015\012WHERE (((prevtblTransac"
    "tion.trnYear) In (SELECT yearID FROM prevtblControl)) AND ((prevtblAccount.accTy"
    "pe)=2))\015\012GROUP BY prevtblAccount.accNo, prevtblAccount.accOpenBalance, pre"
    "vtblAccount.caoID, prevtblAccount.accName;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="prevtblAccount.accName"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
End
