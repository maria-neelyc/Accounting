Operation =1
Option =0
Having ="(((tblTransactionSub.accID)=129))"
Begin InputTables
    Name ="tblTransaction"
    Name ="tblTransactionSub"
End
Begin OutputColumns
    Expression ="tblTransaction.persID"
    Alias ="Bal"
    Expression ="Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount]))"
    Expression ="tblTransactionSub.accID"
    Expression ="tblTransaction.trnYear"
End
Begin Joins
    LeftTable ="tblTransaction"
    RightTable ="tblTransactionSub"
    Expression ="tblTransaction.trnID = tblTransactionSub.trnID"
    Flag =1
End
Begin Groups
    Expression ="tblTransaction.persID"
    GroupLevel =0
    Expression ="tblTransactionSub.accID"
    GroupLevel =0
    Expression ="tblTransaction.trnYear"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =-14
    Top =111
    Right =1866
    Bottom =631
    Left =-1
    Top =-1
    Right =1863
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =155
        Top =3
        Name ="tblTransaction"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =396
        Bottom =155
        Top =0
        Name ="tblTransactionSub"
        Name =""
    End
End
