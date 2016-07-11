Operation =1
Option =0
Begin InputTables
    Name ="qryStatTrialBal1"
    Name ="qryChartWithAccounts"
End
Begin OutputColumns
    Expression ="qryChartWithAccounts.coaRef1"
    Alias ="coaRef2"
    Expression ="qryChartWithAccounts.coaRef2"
    Alias ="coaRef3"
    Expression ="qryChartWithAccounts.coaRef3"
    Alias ="coaRef4"
    Expression ="qryChartWithAccounts.coaRef4"
    Alias ="coaRef5"
    Expression ="qryChartWithAccounts.coaRef5"
    Alias ="Bal"
    Expression ="IIf(Sum(qryStatTrialBal1.Bal)<>0,Sum(qryStatTrialBal1.Bal),Sum(qryChartWithAccou"
        "nts.accOpenBalance))"
    Alias ="OpenBal"
    Expression ="Sum(qryChartWithAccounts.accOpenBalance)"
    Alias ="CR"
    Expression ="Sum(qryStatTrialBal1.CR)"
    Alias ="DR"
    Expression ="Sum(qryStatTrialBal1.DR)"
    Alias ="accID"
    Expression ="qryChartWithAccounts.accID"
    Alias ="accNo"
    Expression ="qryChartWithAccounts.accNo"
    Alias ="accName"
    Expression ="qryChartWithAccounts.accName"
End
Begin Joins
    LeftTable ="qryStatTrialBal1"
    RightTable ="qryChartWithAccounts"
    Expression ="qryStatTrialBal1.caoID=qryChartWithAccounts.coaID"
    Flag =3
End
Begin OrderBy
    Expression ="qryChartWithAccounts.coaRef1"
    Flag =0
    Expression ="qryChartWithAccounts.coaRef2"
    Flag =0
    Expression ="qryChartWithAccounts.coaRef3"
    Flag =0
    Expression ="qryChartWithAccounts.coaRef4"
    Flag =0
    Expression ="qryChartWithAccounts.coaRef5"
    Flag =0
End
Begin Groups
    Expression ="qryChartWithAccounts.coaRef1"
    GroupLevel =0
    Expression ="qryChartWithAccounts.coaRef2"
    GroupLevel =0
    Expression ="qryChartWithAccounts.coaRef3"
    GroupLevel =0
    Expression ="qryChartWithAccounts.coaRef4"
    GroupLevel =0
    Expression ="qryChartWithAccounts.coaRef5"
    GroupLevel =0
    Expression ="qryChartWithAccounts.accID"
    GroupLevel =0
    Expression ="qryChartWithAccounts.accNo"
    GroupLevel =0
    Expression ="qryChartWithAccounts.accName"
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
        dbText "Name" ="accNo"
    End
    Begin
        dbText "Name" ="accName"
    End
    Begin
        dbText "Name" ="coaRef2"
    End
    Begin
        dbText "Name" ="coaRef3"
    End
    Begin
        dbText "Name" ="coaRef4"
    End
    Begin
        dbText "Name" ="coaRef5"
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
    Left =59
    Top =110
    Right =1111
    Bottom =422
    Left =-1
    Top =-1
    Right =1045
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="qryStatTrialBal1"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =14
        Name ="qryChartWithAccounts"
        Name =""
    End
End
