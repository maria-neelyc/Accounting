dbMemo "SQL" ="SELECT tblTransaction.trnID, tblTransaction.trnInternalRef, tblTransaction.trnEn"
    "tryDate, tblTransactionSub.trnsID, tblReference.refName, tblDescription.desName,"
    " tblTransactionSub.trnsNote, tblTransactionSub.trnsDebits, tblTransactionSub.trn"
    "sCredits, tblVat.vatName, (tblVat.vatRate & \"%\") AS vatRate\015\012FROM tblTra"
    "nsaction INNER JOIN (((tblTransactionSub LEFT JOIN tblDescription ON tblTransact"
    "ionSub.desID=tblDescription.desID) LEFT JOIN tblReference ON tblTransactionSub.r"
    "efID=tblReference.refID) LEFT JOIN tblVat ON tblTransactionSub.trnsVID=tblVat.va"
    "tID) ON tblTransaction.trnID=tblTransactionSub.trnID\015\012WHERE (((tblTransact"
    "ionSub.trnsIsVAT)=False));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="vatRate"
    End
End
