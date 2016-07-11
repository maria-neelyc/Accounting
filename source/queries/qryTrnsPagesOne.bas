Operation =1
Option =0
Where ="(((tblTransaction.trnPageCounter)=[Forms].[frmTransaction].[trnPageCounter]))"
Begin InputTables
    Name ="tblTransaction"
    Name ="tblAccount"
    Name ="tblVat"
    Name ="tblTransactionSub"
    Name ="tblReference"
    Name ="tblCurrency"
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
    LeftTable ="tblVat"
    RightTable ="tblTransactionSub"
    Expression ="tblVat.vatID = tblTransactionSub.trnsVAT"
    Flag =3
    LeftTable ="tblTransactionSub"
    RightTable ="tblReference"
    Expression ="tblTransactionSub.refID = tblReference.refID"
    Flag =2
    LeftTable ="tblAccount"
    RightTable ="tblTransactionSub"
    Expression ="tblAccount.accID = tblTransactionSub.accID"
    Flag =1
    LeftTable ="tblAccount"
    RightTable ="tblTransactionSub"
    Expression ="tblAccount.accID = tblTransactionSub.accID"
    Flag =1
    LeftTable ="tblTransaction"
    RightTable ="tblTransactionSub"
    Expression ="tblTransaction.trnID = tblTransactionSub.trnID"
    Flag =1
    LeftTable ="tblTransactionSub"
    RightTable ="tblCurrency"
    Expression ="tblTransactionSub.crnID = tblCurrency.crnID"
    Flag =2
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
        Bottom =113
        Top =0
        Name ="tblTransaction"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="tblAccount"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="tblVat"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =233
        Top =0
        Name ="tblTransactionSub"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =670
        Bottom =113
        Top =0
        Name ="tblReference"
        Name =""
    End
    Begin
        Left =741
        Top =123
        Right =837
        Bottom =230
        Top =0
        Name ="tblCurrency"
        Name =""
    End
End
