Operation =1
Option =0
Where ="(((qryChartWithAccountsVat.accType) In (2,3)))"
Begin InputTables
    Name ="qryVatBal1"
    Name ="qryChartWithAccountsVat"
End
Begin OutputColumns
    Expression ="qryChartWithAccountsVat.coaRef1"
    Expression ="qryChartWithAccountsVat.coaRef2"
    Expression ="qryChartWithAccountsVat.coaRef3"
    Expression ="qryChartWithAccountsVat.coaRef4"
    Expression ="qryChartWithAccountsVat.coaRef5"
    Alias ="Bal"
    Expression ="IIf(IsNumeric(Sum(qryVatBal1.Bal)),Sum(qryVatBal1.Bal),0)"
    Alias ="OpenBal"
    Expression ="IIf(Forms.frmReport.txtFromStr1=0,Sum(qryChartWithAccountsVat.accOpenBalance),0)"
    Alias ="CR"
    Expression ="Sum(qryVatBal1.CR)"
    Alias ="DR"
    Expression ="Sum(qryVatBal1.DR)"
    Alias ="V_CR"
    Expression ="Sum(qryVatBal1.VAmount_Cr)"
    Alias ="V_DR"
    Expression ="Sum(qryVatBal1.VAmount_Dr)"
    Alias ="accID"
    Expression ="IIf(Forms.frmReport.txtToStr2=5,qryChartWithAccountsVat.accID)"
    Expression ="qryChartWithAccountsVat.accNo"
    Expression ="qryChartWithAccountsVat.accName"
    Expression ="qryChartWithAccountsVat.accType"
    Expression ="qryChartWithAccountsVat.VatType"
End
Begin Joins
    LeftTable ="qryVatBal1"
    RightTable ="qryChartWithAccountsVat"
    Expression ="qryVatBal1.caoID = qryChartWithAccountsVat.coaID"
    Flag =3
End
Begin OrderBy
    Expression ="qryChartWithAccountsVat.coaRef1"
    Flag =0
    Expression ="qryChartWithAccountsVat.coaRef2"
    Flag =0
    Expression ="qryChartWithAccountsVat.coaRef3"
    Flag =0
    Expression ="qryChartWithAccountsVat.coaRef4"
    Flag =0
    Expression ="qryChartWithAccountsVat.coaRef5"
    Flag =0
End
Begin Groups
    Expression ="qryChartWithAccountsVat.coaRef1"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.coaRef2"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.coaRef3"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.coaRef4"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.coaRef5"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.accNo"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.accName"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.accType"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.VatType"
    GroupLevel =0
    Expression ="qryChartWithAccountsVat.accID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Bal"
    End
    Begin
        dbText "Name" ="accID"
    End
    Begin
        dbText "Name" ="CR"
    End
    Begin
        dbText "Name" ="DR"
    End
    Begin
        dbText "Name" ="OpenBal"
    End
End
Begin
    State =0
    Left =44
    Top =57
    Right =1096
    Bottom =394
    Left =-1
    Top =-1
    Right =1035
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =236
        Bottom =106
        Top =4
        Name ="qryVatBal1"
        Name =""
    End
    Begin
        Left =367
        Top =-1
        Right =528
        Bottom =111
        Top =15
        Name ="qryChartWithAccountsVat"
        Name =""
    End
End
