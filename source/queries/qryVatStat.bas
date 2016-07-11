Operation =1
Option =2
Begin InputTables
    Name ="qryVatAcc"
    Name ="qryStatementAllBal"
End
Begin OutputColumns
    Expression ="qryStatementAllBal.*"
    Expression ="qryVatAcc.VatType"
End
Begin Joins
    LeftTable ="qryVatAcc"
    RightTable ="qryStatementAllBal"
    Expression ="qryVatAcc.accNo = qryStatementAllBal.accNo"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
Begin
End
Begin
    State =0
    Left =40
    Top =94
    Right =1258
    Bottom =649
    Left =-1
    Top =-1
    Right =1201
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =275
        Bottom =193
        Top =0
        Name ="qryVatAcc"
        Name =""
    End
    Begin
        Left =345
        Top =-3
        Right =726
        Bottom =240
        Top =0
        Name ="qryStatementAllBal"
        Name =""
    End
End
