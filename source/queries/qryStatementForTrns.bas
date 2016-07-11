Operation =1
Option =0
Where ="(((tblTransaction.persID) Between Forms.frmReport.txtFromStr2 And Forms.frmRepor"
    "t.txtToStr2))"
Begin InputTables
    Name ="tblControl"
    Name ="tblTransaction"
    Name ="tblTransactionSub"
    Name ="tblReference"
    Name ="tblDescription"
End
Begin OutputColumns
    Expression ="tblTransactionSub.trnsID"
    Expression ="tblTransaction.trnEntryDate"
    Expression ="tblControl.yearID"
    Expression ="tblTransaction.persID"
    Expression ="tblTransactionSub.trnsNote"
    Alias ="trnsDebits"
    Expression ="IIf(tblTransactionSub.trnsSign=\"D\",tblTransactionSub.trnsFAmount,0)"
    Alias ="trnsCredits"
    Expression ="IIf(tblTransactionSub.trnsSign=\"C\",tblTransactionSub.trnsFAmount,0)"
    Expression ="tblTransaction.trnInternalRef"
    Expression ="tblReference.refName"
    Expression ="tblDescription.desName"
    Expression ="tblTransactionSub.trnsDocDate"
    Expression ="tblTransactionSub.trnsDate"
    Expression ="tblTransactionSub.accID"
    Expression ="tblTransactionSub.docNo"
End
Begin Joins
    LeftTable ="tblControl"
    RightTable ="tblTransaction"
    Expression ="tblControl.yearID=tblTransaction.trnYear"
    Flag =1
    LeftTable ="tblTransactionSub"
    RightTable ="tblReference"
    Expression ="tblTransactionSub.refID=tblReference.refID"
    Flag =2
    LeftTable ="tblTransactionSub"
    RightTable ="tblDescription"
    Expression ="tblTransactionSub.desID=tblDescription.desID"
    Flag =2
    LeftTable ="tblTransaction"
    RightTable ="tblTransactionSub"
    Expression ="tblTransaction.trnID=tblTransactionSub.trnID"
    Flag =1
End
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
Begin
    State =0
    Left =2
    Top =78
    Right =1198
    Bottom =390
    Left =-1
    Top =-1
    Right =1185
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =98
        Top =0
        Name ="tblControl"
        Name =""
    End
    Begin
        Left =163
        Top =6
        Right =291
        Bottom =98
        Top =0
        Name ="tblTransaction"
        Name =""
    End
    Begin
        Left =332
        Top =4
        Right =482
        Bottom =96
        Top =0
        Name ="tblTransactionSub"
        Name =""
    End
    Begin
        Left =499
        Top =6
        Right =595
        Bottom =83
        Top =0
        Name ="tblReference"
        Name =""
    End
    Begin
        Left =651
        Top =6
        Right =747
        Bottom =83
        Top =0
        Name ="tblDescription"
        Name =""
    End
End
