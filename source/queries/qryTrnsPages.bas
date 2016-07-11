Operation =1
Option =0
Where ="(((tblTransaction.trnPageCounter)>=[Forms].[frmReport].[txtFromStr1] And (tblTra"
    "nsaction.trnPageCounter)<=[Forms].[frmReport].[txtToStr1]) AND ((tblTransaction."
    "persID)>=[Forms].[frmReport].[txtFromStr2] And (tblTransaction.persID)<=[Forms]."
    "[frmReport].[txtToStr2]))"
Begin InputTables
    Name ="tblTransaction"
    Name ="tblAccount"
    Name ="tblTransactionSub"
    Name ="tblReference"
    Name ="tblCurrency"
    Name ="tblVat"
End
Begin OutputColumns
    Expression ="tblTransaction.trnPageCounter"
    Expression ="tblTransaction.trnInternalRef"
    Expression ="tblTransaction.trnEntryDate"
    Expression ="tblTransaction.persID"
    Expression ="tblTransaction.trnYear"
    Expression ="tblTransaction.trnLock"
    Expression ="tblAccount.accNo"
    Expression ="tblAccount.accName"
    Expression ="tblTransactionSub.trnsDate"
    Expression ="tblVat.vatName"
    Expression ="tblReference.refShortName"
    Expression ="tblTransactionSub.docNo"
    Expression ="tblTransactionSub.trnsNote"
    Expression ="tblTransactionSub.trnsDebits"
    Expression ="tblTransactionSub.trnsCredits"
    Expression ="tblTransactionSub.trnsDocDate"
    Alias ="FCredits"
    Expression ="tblTransactionSub.trnsCredits*tblTransactionSub.trnsRate"
    Alias ="FDebits"
    Expression ="tblTransactionSub.trnsDebits*tblTransactionSub.trnsRate"
    Expression ="tblCurrency.crnShortName"
End
Begin Joins
    LeftTable ="tblTransactionSub"
    RightTable ="tblReference"
    Expression ="tblTransactionSub.refID = tblReference.refID"
    Flag =2
    LeftTable ="tblTransactionSub"
    RightTable ="tblVat"
    Expression ="tblTransactionSub.trnsVID = tblVat.vatID"
    Flag =2
    LeftTable ="tblAccount"
    RightTable ="tblTransactionSub"
    Expression ="tblAccount.accID = tblTransactionSub.accID"
    Flag =1
    LeftTable ="tblTransactionSub"
    RightTable ="tblCurrency"
    Expression ="tblTransactionSub.crnID = tblCurrency.crnID"
    Flag =2
    LeftTable ="tblAccount"
    RightTable ="tblTransactionSub"
    Expression ="tblAccount.accID = tblTransactionSub.accID"
    Flag =1
    LeftTable ="tblTransaction"
    RightTable ="tblTransactionSub"
    Expression ="tblTransaction.trnID = tblTransactionSub.trnID"
    Flag =1
End
Begin OrderBy
    Expression ="tblTransaction.trnPageCounter"
    Flag =0
    Expression ="tblTransaction.trnInternalRef"
    Flag =0
    Expression ="tblTransactionSub.trnsDate"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="FCredits"
    End
    Begin
        dbText "Name" ="FDebits"
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1258
    Bottom =495
    Left =-1
    Top =-1
    Right =1233
    Bottom =287
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =158
        Top =0
        Name ="tblTransaction"
        Name =""
    End
    Begin
        Left =738
        Top =1
        Right =834
        Bottom =108
        Top =8
        Name ="tblAccount"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =361
        Bottom =248
        Top =2
        Name ="tblTransactionSub"
        Name =""
    End
    Begin
        Left =449
        Top =163
        Right =545
        Bottom =255
        Top =0
        Name ="tblReference"
        Name =""
    End
    Begin
        Left =750
        Top =129
        Right =846
        Bottom =236
        Top =0
        Name ="tblCurrency"
        Name =""
    End
    Begin
        Left =609
        Top =44
        Right =705
        Bottom =151
        Top =0
        Name ="tblVat"
        Name =""
    End
End
