Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    TabularCharSet =128
    TabularFamily =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =9319
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =3516
    Top =3240
    Right =12768
    Bottom =10692
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa05aa56390cce340
    End
    RecordSource ="SELECT tblAccount.*, appCustomer.custID, appCustomer.custName, appCustomer.custS"
        "Name, appCustomer.custAddress, appCustomer.custPhone1, appCustomer.custPhone2, a"
        "ppCustomer.custMobile, appCustomer.custFax, appCustomer.custEmail, appCustomer.c"
        "ustTown, appCustomer.custCountry, appCustomer.custPOBox, appCustomer.custBankAcc"
        ", appCustomer.custBankAccNo, appCustomer.custIBAN FROM tblAccount LEFT JOIN appC"
        "ustomer ON tblAccount.accID = appCustomer.accID;"
    Caption =" PROMETHEUS ( 2011 )"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName =""
    PrtMip = Begin
        0xae050000ae050000ae050000ae05000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontCharSet =238
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =238
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =238
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Tab
            TextFontCharSet =238
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =1243
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    Left =6009
                    Top =868
                    Width =1920
                    Height =300
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="lblPos"
                    Caption ="Rec. New"
                    FontName ="MS Sans Serif"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =161
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5670
                    Left =414
                    Top =868
                    Width =2415
                    Height =300
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboFind"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount OR"
                        "DER BY tblAccount.accName; "
                    ColumnWidths ="0;3402;1134"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    FontName ="Arial Greek"
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =182
                    Top =283
                    TabIndex =1
                    Name ="HasChanged"
                    DefaultValue ="0"
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =90
                    Top =525
                    Width =3969
                    Height =703
                    TabIndex =2
                    Name ="grpFind"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextFontCharSet =161
                    ColumnCount =2
                    ListWidth =2880
                    Left =1413
                    Top =584
                    Width =1418
                    Height =260
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cboSearch"
                    RowSourceType ="Value List"
                    RowSource ="\"Account Name\";1;\"Account Code\";2"
                    ColumnWidths ="1444;0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="\"Account Name\""
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =397
                    Top =584
                    Width =960
                    Height =225
                    FontWeight =700
                    BackColor =-2147483633
                    Name ="Label150"
                    Caption ="Search By"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =247
                    Left =354
                    Top =23
                    Width =2640
                    Height =330
                    FontSize =14
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Accounts"
                    FontName ="Microsoft Sans Serif"
                End
            End
        End
        Begin Section
            Height =5433
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    BackStyle =0
                    Left =45
                    Top =90
                    Width =8310
                    Height =5250
                    Name ="tabAcc"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =156
                            Top =456
                            Width =8088
                            Height =4776
                            Name ="pgeAcc"
                            Caption ="Account Details"
                            Begin
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =7883
                                    Top =575
                                    Width =180
                                    Height =255
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accID"
                                    ControlSource ="accID"
                                    FontName ="MS Sans Serif"
                                End
                                Begin OptionGroup
                                    OverlapFlags =223
                                    Left =435
                                    Top =645
                                    Width =3900
                                    Height =1240
                                    TabIndex =1
                                    BorderColor =8421504
                                    Name ="frmRef"
                                    DefaultValue ="1"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Enabled = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =11
                                    ListWidth =3600
                                    Left =1275
                                    Top =765
                                    Width =2970
                                    Height =270
                                    TabIndex =2
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";"
                                        "\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboSubAccount"
                                    ControlSource ="caoID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblChart.coaID, tblChart.coaName, tblChart.coaNo, tblChart.coaIDg, tblCha"
                                        "rt.lvlID, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4"
                                        ", tblChart.coaRef5 FROM tblChart WHERE (((tblChart.coaID) Not In (SELECT tblAcco"
                                        "unt.caoID FROM tblAccount)) AND ((tblChart.lvlID)=5)) ORDER BY tblChart.coaName,"
                                        " tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChar"
                                        "t.coaRef5; "
                                    ColumnWidths ="0;2886;0;0;0;0;0;0;0;0;0"
                                    AfterUpdate ="[Event Procedure]"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    OnDblClick ="[Event Procedure]"
                                    FontName ="Arial Greek"
                                    OnClick ="[Event Procedure]"
                                    OnChange ="[Event Procedure]"
                                End
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =1905
                                    Top =2040
                                    Width =3330
                                    Height =255
                                    TabIndex =3
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accNo"
                                    ControlSource ="accNo"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Arial Greek"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =465
                                            Top =2040
                                            Width =1380
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="bnkaABA_Label"
                                            Caption ="Account Number"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =1905
                                    Top =2370
                                    Width =3330
                                    Height =255
                                    TabIndex =4
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accName1"
                                    ControlSource ="accName"
                                    FontName ="Arial Greek"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =465
                                            Top =2370
                                            Width =1380
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="bnkaBLZ_Label"
                                            Caption ="Account Name"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =1954
                                    Top =2755
                                    TabIndex =5
                                    Name ="accStatus"
                                    ControlSource ="accStatus"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =652
                                            Top =2708
                                            Width =1200
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label33"
                                            Caption ="Active"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =3861
                                    Top =2760
                                    TabIndex =6
                                    Name ="accIsVat"
                                    ControlSource ="accIsVat"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2563
                                            Top =2708
                                            Width =1170
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label31"
                                            Caption ="V.A.T. Account"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =5279
                                    Top =2760
                                    TabIndex =7
                                    Name ="accVatable"
                                    ControlSource ="accVatable"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =4288
                                            Top =2708
                                            Width =885
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label83"
                                            Caption ="Vat-able"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =6811
                                    Top =2752
                                    TabIndex =8
                                    Name ="accEasyDoc"
                                    ControlSource ="accEasyDoc"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =5820
                                            Top =2708
                                            Width =885
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label166"
                                            Caption ="EasyDoc"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =1191
                                    Left =1898
                                    Top =3720
                                    Width =1725
                                    Height =270
                                    TabIndex =9
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accSign"
                                    ControlSource ="accSign"
                                    RowSourceType ="Value List"
                                    RowSource ="C;Credit;D;Debit"
                                    ColumnWidths ="0;1134"
                                    AfterUpdate ="[Event Procedure]"
                                    OnEnter ="[Event Procedure]"
                                    FontName ="Arial Greek"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =1058
                                            Top =3720
                                            Width =780
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="crnID_Label"
                                            Caption ="Sign"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =4545
                                    Top =645
                                    Width =3667
                                    Height =1249
                                    Name ="Box130"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =4638
                                    Top =750
                                    Width =645
                                    Height =240
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="lblLev1"
                                    Caption ="Level 1:"
                                    FontName ="MS Sans Serif"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =1134
                                    Left =1913
                                    Top =3030
                                    Width =1740
                                    Height =270
                                    TabIndex =10
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accType"
                                    ControlSource ="accType"
                                    RowSourceType ="Value List"
                                    RowSource ="1;Profit & Loss;2;Balance Sheet"
                                    ColumnWidths ="0;1134"
                                    AfterUpdate ="[Event Procedure]"
                                    OnEnter ="[Event Procedure]"
                                    DefaultValue ="2"
                                    FontName ="Arial Greek"
                                    OnChange ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =1073
                                            Top =3030
                                            Width =780
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label35"
                                            Caption ="Type"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    TextFontCharSet =161
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =1920
                                    Top =3388
                                    Width =1710
                                    Height =255
                                    TabIndex =11
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accOpenBalance"
                                    ControlSource ="accOpenBalance"
                                    Format ="Standard"
                                    FontName ="Arial Greek"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =660
                                            Top =3388
                                            Width =1200
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="bnkaIBAN_Label"
                                            Caption ="Open Balance"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =5352
                                    Top =750
                                    Width =2799
                                    TabIndex =12
                                    BackColor =-2147483643
                                    BorderColor =8421504
                                    Name ="txtLev1"
                                    FontName ="MS Sans Serif"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =4635
                                    Top =1020
                                    Width =645
                                    Height =240
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="lblLvl2"
                                    Caption ="Level 2:"
                                    FontName ="MS Sans Serif"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =4635
                                    Top =1290
                                    Width =645
                                    Height =240
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="lblLvl3"
                                    Caption ="Level 3:"
                                    FontName ="MS Sans Serif"
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =5349
                                    Top =1575
                                    Width =2799
                                    TabIndex =13
                                    BackColor =-2147483643
                                    BorderColor =8421504
                                    Name ="txtLev4"
                                    FontName ="MS Sans Serif"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =4635
                                    Top =1575
                                    Width =645
                                    Height =240
                                    BackColor =-2147483633
                                    ForeColor =-2147483630
                                    Name ="lblLev4"
                                    Caption ="Level 4:"
                                    FontName ="MS Sans Serif"
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =540
                                    Top =765
                                    Width =534
                                    Height =271
                                    TabIndex =14
                                    BackColor =-2147483643
                                    BorderColor =8421504
                                    Name ="txtLev5"
                                    ControlSource ="=Right(\"00\" & CStr(IIf(IsNull([cboSubAccount].[column](9)),0,[cboSubAccount].["
                                        "column](9))),3)"
                                    FontName ="MS Sans Serif"
                                End
                                Begin Rectangle
                                    SpecialEffect =1
                                    OverlapFlags =223
                                    Left =605
                                    Top =1166
                                    Width =3591
                                    Height =590
                                    Name ="Box145"
                                End
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =652
                                    Top =1213
                                    Width =3424
                                    Height =474
                                    BackColor =-2147483633
                                    ForeColor =128
                                    Name ="lblPrvRef"
                                    Caption ="Your previous reference was:"
                                    FontName ="MS Sans Serif"
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =5349
                                    Top =1020
                                    Width =2799
                                    TabIndex =15
                                    BackColor =-2147483643
                                    BorderColor =8421504
                                    Name ="txtLev2"
                                    FontName ="MS Sans Serif"
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =5349
                                    Top =1290
                                    Width =2799
                                    TabIndex =16
                                    BackColor =-2147483643
                                    BorderColor =8421504
                                    Name ="txtLev3"
                                    FontName ="MS Sans Serif"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =156
                            Top =456
                            Width =8088
                            Height =4776
                            Name ="pgeBalance"
                            Caption ="Current Balance"
                            Begin
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =551
                                    Width =1440
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accOpBal"
                                    ControlSource ="accOpenBalance"
                                    ValidationText ="Has to be a value"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    InputMask ="##########.##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =700
                                            Top =551
                                            Width =1335
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblOpBal"
                                            Caption ="Opening Ballance"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =791
                                    Width =1440
                                    TabIndex =1
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer1"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =791
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblper1"
                                            Caption ="1"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =1031
                                    Width =1440
                                    TabIndex =2
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer2"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =1031
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer2"
                                            Caption ="2"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =1271
                                    Width =1440
                                    TabIndex =3
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer3"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =1271
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer3"
                                            Caption ="3"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =1511
                                    Width =1440
                                    TabIndex =4
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer4"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =1511
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer4"
                                            Caption ="4"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =1751
                                    Width =1440
                                    TabIndex =5
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer5"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =1751
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer5"
                                            Caption ="5"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =1991
                                    Width =1440
                                    TabIndex =6
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer6"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =1991
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer6"
                                            Caption ="6"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2161
                                    Top =2231
                                    Width =1440
                                    TabIndex =7
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer7"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =720
                                            Top =2231
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer7"
                                            Caption ="7"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2161
                                    Top =2471
                                    Width =1440
                                    TabIndex =8
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer8"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =720
                                            Top =2471
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer8"
                                            Caption ="8"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =2711
                                    Width =1440
                                    TabIndex =9
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer9"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =2711
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer9"
                                            Caption ="9"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =2951
                                    Width =1440
                                    TabIndex =10
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer10"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =2951
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer10"
                                            Caption ="10"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =3191
                                    Width =1440
                                    TabIndex =11
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer11"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =3191
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer11"
                                            Caption ="11"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2164
                                    Top =3431
                                    Width =1440
                                    TabIndex =12
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accPer12"
                                    Format ="Standard"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =723
                                            Top =3431
                                            Width =615
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblPer12"
                                            Caption ="12"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4680
                                    Top =555
                                    Width =1440
                                    TabIndex =13
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="accOpenBalCur"
                                    ControlSource ="accOpenBalCur"
                                    ValidationText ="Has to be a value"
                                    DefaultValue ="0"
                                    FontName ="MS Sans Serif"
                                    InputMask ="##########.##"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =3675
                                    Top =555
                                    Width =946
                                    TabIndex =14
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"6\""
                                    Name ="accCur"
                                    ControlSource ="accCur"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblCurrency.crnID, tblCurrency.crnShortName FROM tblCurrency; "
                                    ColumnWidths ="0;567"
                                    FontName ="MS Sans Serif"
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =580
                                    Width =2475
                                    Height =255
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custName"
                                    ControlSource ="custName"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =575
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =13209
                                            Name ="agName Label"
                                            Caption ="Name"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agName_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =920
                                    Width =2475
                                    Height =255
                                    TabIndex =1
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custSName"
                                    ControlSource ="custSName"
                                    StatusBarText ="Short Name"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =915
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agSName Label"
                                            Caption ="Short Name"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agSName_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    FELineBreak = NotDefault
                                    ScrollBars =2
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =1260
                                    Width =3339
                                    Height =915
                                    TabIndex =2
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custAddress"
                                    ControlSource ="custAddress"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =1255
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agAddress Label"
                                            Caption ="Address"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agAddress_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =2245
                                    Width =2490
                                    Height =255
                                    TabIndex =3
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custTown"
                                    ControlSource ="custTown"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =2275
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agTown Label"
                                            Caption ="Town"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agTown_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1458
                                    Top =2980
                                    Width =1185
                                    Height =255
                                    TabIndex =4
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custPOBox"
                                    ControlSource ="custPOBox"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =3012
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agPOBox Label"
                                            Caption ="P.O. Box"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agPOBox_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1458
                                    Top =3322
                                    Width =1185
                                    Height =255
                                    TabIndex =5
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custPhone1"
                                    ControlSource ="custPhone1"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =3353
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agPhone1 Label"
                                            Caption ="Phones"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agPhone1_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =2753
                                    Top =3305
                                    Width =1185
                                    Height =255
                                    TabIndex =6
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custPhone2"
                                    ControlSource ="custPhone2"
                                    FontName ="MS Sans Serif"
                                    AsianLineBreak =0
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =3928
                                    Width =1185
                                    Height =255
                                    TabIndex =7
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custFax"
                                    ControlSource ="custFax"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =3920
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agFax1 Label"
                                            Caption ="Faxes"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agFax1_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =1457
                                    Top =4268
                                    Width =3231
                                    TabIndex =8
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custEmail"
                                    ControlSource ="custEmail"
                                    StatusBarText ="Email"
                                    FontName ="MS Sans Serif"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =4260
                                            Width =675
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label127"
                                            Caption ="Email"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =3628
                                    Width =1185
                                    Height =255
                                    TabIndex =9
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custMobile"
                                    ControlSource ="custMobile"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =3636
                                            Width =912
                                            Height =255
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="Label134"
                                            Caption ="Mobile"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =1457
                                    Top =2575
                                    Width =2490
                                    Height =255
                                    TabIndex =10
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custCountry"
                                    ControlSource ="custCountry"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =2616
                                            Width =915
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="agCountry Label"
                                            Caption ="Country"
                                            FontName ="MS Sans Serif"
                                            EventProcPrefix ="agCountry_Label"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =1447
                                    Top =4677
                                    TabIndex =11
                                    Name ="custBankAcc"
                                    ControlSource ="custBankAcc"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =260
                                            Top =4600
                                            Width =1125
                                            Height =240
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblBankAccChkb"
                                            Caption ="Bank Account"
                                            FontName ="MS Sans Serif"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =1730
                                    Top =4615
                                    Width =3231
                                    TabIndex =12
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custBankAccNo"
                                    ControlSource ="custBankAccNo"
                                    StatusBarText ="Email"
                                    FontName ="MS Sans Serif"
                                    AsianLineBreak =0
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =161
                                    Left =5147
                                    Top =622
                                    Width =555
                                    Height =255
                                    TabIndex =13
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="custID"
                                    ControlSource ="custID"
                                    FontName ="Arial Greek"
                                    AsianLineBreak =0
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =779
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    TextFontCharSet =0
                    Left =75
                    Top =165
                    Width =366
                    Height =360
                    ForeColor =0
                    Name ="cmdFirst"
                    Caption ="&m"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada44dadad1dadaadad44adad11adaddada44dad111dada ,
                        0xadad44ad1111adaddada44d11111dadaadad44ad1111adaddada44dad111dada ,
                        0xadad44adad11adaddada44dadad1dadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="First Record (Alt+M)"
                    UnicodeAccessKey =109
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =47
                    TextFontCharSet =0
                    Left =1335
                    Top =165
                    Width =366
                    Height =360
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdLast"
                    Caption ="&/"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadad1dadad44adaadada11dada44daddadad111dad44ada ,
                        0xadada1111da44daddadad11111d44adaadada1111da44daddadad111dad44ada ,
                        0xadada11dada44daddadad1dadad44adaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Last Record (Alt+/)"
                    UnicodeAccessKey =47
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =46
                    TextFontCharSet =0
                    Left =915
                    Top =165
                    Width =366
                    Height =360
                    TabIndex =2
                    ForeColor =0
                    Name ="cmdNext"
                    Caption ="&."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Next Record (Alt+.)"
                    UnicodeAccessKey =46
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =44
                    TextFontCharSet =0
                    Left =495
                    Top =165
                    Width =366
                    Height =360
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdPrevious"
                    Caption ="&,"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Previous Record (Alt+,)"
                    UnicodeAccessKey =44
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    TextFontCharSet =0
                    Left =1875
                    Top =165
                    Width =967
                    Height =360
                    TabIndex =4
                    ForeColor =0
                    Name ="cmdAdd"
                    Caption ="Add &New"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Add New Record  (ALT+N)"
                    UnicodeAccessKey =78
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontCharSet =0
                    Left =6223
                    Top =165
                    Width =967
                    Height =360
                    TabIndex =5
                    ForeColor =0
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Delete Record (ALT+D)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    TextFontCharSet =0
                    Left =7280
                    Top =165
                    Width =967
                    Height =360
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Close Current Form (ALT+X)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    TextFontCharSet =0
                    Left =5136
                    Top =165
                    Width =967
                    Height =360
                    TabIndex =7
                    ForeColor =0
                    Name ="cmdEdit"
                    Caption ="&Edit"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Edit Record (ALT+E)"
                    UnicodeAccessKey =69
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =83
                    TextFontCharSet =0
                    Left =4049
                    Top =165
                    Width =967
                    Height =360
                    TabIndex =8
                    ForeColor =0
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Save Record (ALT+S)"
                    UnicodeAccessKey =83
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =85
                    TextFontCharSet =0
                    Left =2962
                    Top =165
                    Width =967
                    Height =360
                    TabIndex =9
                    ForeColor =0
                    Name ="cmdUndo"
                    Caption ="&Undo"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Undo Last Change (ALT+U)"
                    UnicodeAccessKey =85
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontCharSet =0
                    Left =8352
                    Top =168
                    Width =967
                    Height =360
                    TabIndex =10
                    ForeColor =0
                    Name ="cmdChart"
                    Caption ="&Chart"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Show Chart of Account (ALT+C)"
                    UnicodeAccessKey =67
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public OldValue As String
Public OldId As Long
Public LastID As Long
Private Sub accEasyDoc_Click()
    If (accEasyDoc = 0) Then
    'delete customers that have empty name field and assignment to this account
        'check first if there are any documents connected with this account id
        Dim rs As Recordset
        strSQL = "SELECT appInvoice.InvID " & _
                 "FROM appCustomer INNER JOIN appInvoice ON appCustomer.custID = appInvoice.custID " & _
                 "WHERE appCustomer.accID = " & accID
        FnLog (strSQL)
        Set rs = CurrentDb.OpenRecordset(strSQL)
        If (rs.RecordCount > 0) Then
            accEasyDoc.Value = True
            MsgBox ("There is a document associated with this account. You cannot remove is from easy doc.")
            DoCmd.CancelEvent
            Exit Sub
        End If
        
        strSQL = "DELETE FROM appCustomer WHERE accID = " & accID & " AND custName IS NULL"
        DoCmd.SetWarnings False
        FnLog (strSQL)
        DoCmd.RunSQL (strSQL)
        DoCmd.SetWarnings True
    Else
        DoCmd.SetWarnings False
    
    'Updates local control table
    strSQL = "INSERT INTO appCustomer (accID, CustName) " & _
             "VALUES (" & Me.accID & ", '" & Me.accName1 & "')"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
        
         DoCmd.SetWarnings True
    End If
End Sub

Private Sub accIsVat_AfterUpdate()
HasChanged = True
End Sub

Private Sub accName_AfterUpdate()
HasChanged = True
End Sub


Private Sub accNo_AfterUpdate()
HasChanged = True
End Sub
Private Sub accSign_AfterUpdate()
HasChanged = True
End Sub

Private Sub accSign_Enter()
    Me.accSign.Dropdown
End Sub

Private Sub accStatus_AfterUpdate()
HasChanged = True
End Sub

Private Sub accType_AfterUpdate()
HasChanged = True
End Sub

Private Sub accType_Change()
Dim A As Variant
End Sub

Private Sub accType_Enter()
'    Me.accType.Dropdown
End Sub

Private Sub accVatable_AfterUpdate()
HasChanged = True
End Sub

Private Sub cboFind_AfterUpdate()
On Error GoTo Err_End
' Find the record that matches the control.
If IsNoData(cboFind.Value) = False Then

    If Me.AllowEdits = True Then
        If Me.NewRecord Then
            'Me.Undo
            Call cmdUndo_Click
        Else
            DoCmd.RunCommand acCmdSaveRecord
        End If
        
        cmdDel.Enabled = True
        cmdExit.Enabled = True
        cmdEdit.Enabled = True
        cmdAdd.Enabled = True
        cmdUndo.Enabled = False
        cmdSave.Enabled = False
        Me.AllowEdits = False
    End If



    If fsClosedYear And HasChanged Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
        Else
            If IsNull(accNoCopy) Then
                accNoCopy = accNo
            End If
            sUpdateAccount (accNoCopy)
        End If
    End If

    'Me.Requery
    
    Me.RecordsetClone.FindFirst "[accID] = " & Me![cboFind]
    DoEvents
    Me.Bookmark = Me.RecordsetClone.Bookmark
    
End If

Exit Sub
Err_End:
        Dim strMsg As String
        strMsg = FnIsErr(Err.Number)
        If strMsg <> "" Or IsNull(strMsg) Or IsEmpty(strMsg) Then
            MsgBox strMsg, vbExclamation, "Error!"
        Else
        End If
        cboFind.Value = ""
End Sub

Private Sub cboGroup_Change()
If IsNull(cboGroup) = False Then
    Select Case frmRef2
        Case 1
            cboDepartment.ColumnWidths = "0cm;5.087cm;1.27cm;0cm;0cm"
            cboDepartment.RowSource = "SELECT tblChart.coaID, tblChart.coaName, tblChart.coaNo, tblChart.coaIDg, tblChart.lvlID FROM tblChart WHERE (((tblChart.coaIDg)<>[coaID] And (tblChart.coaIDg)=" & cboGroup & "));"
        Case 2
            cboDepartment.ColumnWidths = "0cm;1.27cm;5.087cm;0cm;0cm"
            cboDepartment.RowSource = "SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, tblChart.coaIDg, tblChart.lvlID FROM tblChart WHERE (((tblChart.coaIDg)<>[coaID] And (tblChart.coaIDg)=" & cboGroup & "));"
    End Select
    'lstAccount.RowSource = "SELECT tblChart.coaID, tblChart.coaName, tblChart.coaNo, tblChart.coaIDg, tblChart.lvlID FROM tblChart WHERE (((tblChart.coaIDg)<>[coaID] And (tblChart.coaIDg)=" & cboGroup & "));"
    'cboDepartment.Value = Null
    'cboAccount.Value = Null
End If
End Sub

Private Sub cboFind_Enter()
    Me.AllowEdits = True
    'Me.cboFind.Dropdown
End Sub

Private Sub cboFind_Exit(Cancel As Integer)
    If Me.cmdSave.Enabled = True Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
End Sub

Private Sub cboSearch_AfterUpdate()
Call cmdSave_Click
Select Case cboSearch.Column(1)
    Case 1
      'cboFind.InputMask = ""
      cboFind.ColumnWidths = "0cm;6cm;2cm"
      cboFind.RowSource = "SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo " & _
            " FROM tblAccount ORDER BY tblAccount.accName;"
    Case Else
    cboFind.ColumnWidths = "0cm;2cm;6cm"
      cboFind.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
            " FROM tblAccount ORDER BY tblAccount.accNo;"
    
End Select
End Sub
Private Sub cboSearch_Enter()
    Me.AllowEdits = True
End Sub

Private Sub cboSearch_Exit(Cancel As Integer)
    grpFind.DefaultValue = cboSearch.Column(1)
    
    Me.AllowEdits = False
End Sub
Private Sub cboSubAccount_AfterUpdate()
HasChanged = True
End Sub

Private Sub cboSubAccount_Change()
If OldId <> cboSubAccount.Column(0) And Me.NewRecord = False Then
    lblPrvRef.Caption = "Your previous reference was: " & OldValue
    lblPrvRef.Visible = True
Else
    lblPrvRef.Visible = False
End If
End Sub

Private Sub cboSubAccount_Click()
    Dim rs As Recordset
    Dim strSQL As String

    strSQL = "SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & CStr(tblChart.coaRef1), 2) AS coaRef1, Right(Chr(48) & CStr(tblChart.coaRef2), 2) AS coaRef2, Right(Chr(48) & CStr(tblChart.coaRef3), 2) AS coaRef3, Right(Chr(48) & CStr(tblChart.coaRef4), 2) As coaRef4, tblChart.lvlID FROM tblChart WHERE (tblChart.coaRef1 = " & cboSubAccount.Column(5) & " And tblChart.lvlID = 1) OR (tblChart.coaRef1 & tblChart.coaRef2 =  " & cboSubAccount.Column(5) & cboSubAccount.Column(6) & " And tblChart.lvlID = 2) OR (tblChart.coaRef1 & tblChart.coaRef2& tblChart.coaRef3 =  " & cboSubAccount.Column(5) & cboSubAccount.Column(6) & cboSubAccount.Column(7) & " And tblChart.lvlID = 3) OR (tblChart.coaRef1 & tblChart.coaRef2& tblChart.coaRef3& tblChart.coaRef4 =  " & cboSubAccount.Column(5) & cboSubAccount.Column(6) & cboSubAccount.Column(7) & cboSubAccount.Column(8) & " And tblChart.lvlID = 4) ORDER BY tblChart.lvlID"

    Set rs = CurrentDb.OpenRecordset(strSQL)

    txtLev1 = ""
    txtLev2 = ""
    txtLev3 = ""
    txtLev4 = ""

    

    Do While (Not rs.EOF)
        Select Case rs.Fields("lvlID")
        Case 1
            txtLev1 = Right("0" & rs.Fields("coaRef1"), 2) & "  " & rs.Fields("coaName")
        Case 2
            txtLev2 = Right("0" & rs.Fields("coaRef2"), 2) & "  " & rs.Fields("coaName")
        Case 3
            txtLev3 = Right("0" & rs.Fields("coaRef3"), 2) & "  " & rs.Fields("coaName")
        Case 4
            txtLev4 = Right("0" & rs.Fields("coaRef4"), 2) & "  " & rs.Fields("coaName")
        End Select
        rs.MoveNext
    Loop
        rs.Close

End Sub

Private Sub cboSubAccount_DblClick(Cancel As Integer)
If cmdEdit.Enabled = False Then
    Me.Modal = False
    DoCmd.OpenForm "frmChart", , , , , , Me.Form.name
End If
End Sub

Private Sub cboSubAccount_Enter()
'Me.cboSubAccount.Dropdown
End Sub

Private Sub cboSubAccount_Exit(Cancel As Integer)
Dim coaIDExists As Variant
'coaIDExists = DLookup("[coaID]", "tblAccount", "[coaID] = " & cboSubAccount.Column(0) & " AND [accNo] <> " & Me.accNo)
Dim rs As Recordset

strSQL = "SELECT tblAccount.accNo, tblAccount.accName FROM tblAccount WHERE tblAccount.caoID = " & Nz(caoID, 0) & " AND tblAccount.accNo <> '" & Nz(accNo, 0) & "'"

Set rs = CurrentDb.OpenRecordset(strSQL)
If rs.RecordCount > 0 Then
    MsgBox ("This reference is already assigned to account: " & rs.Fields(0) & " - " & rs.Fields(1))
'    DoCmd.RunCommand acCmdUndo
'    DoCmd.CancelEvent
    Me.cboSubAccount = OldId
    Me.cboSubAccount.Requery
    lblPrvRef.Visible = False
    DoCmd.CancelEvent
    Exit Sub
End If
If IsNull(cboSubAccount.Column(1)) = False Then
   accName1.Value = Left(cboSubAccount.Column(1), 30)
    
    If Me.NewRecord = True Then
        accNo.Value = Right("0" & CStr(cboSubAccount.Column(5)), 2) _
                    & Right("0" & CStr(cboSubAccount.Column(6)), 2) _
                    & Right("0" & CStr(cboSubAccount.Column(7)), 2) _
                    & Right("0" & CStr(cboSubAccount.Column(8)), 2) _
                    & Right("00" & CStr(cboSubAccount.Column(9)), 3)
        accStatus.SetFocus
    End If
End If

End Sub




Private Sub cmdChart_Click()
If cmdEdit.Enabled = True Then
    Me.Modal = False
    DoCmd.OpenForm "frmChart", , , , , , Me.Form.name
End If
End Sub

Private Sub cmdEdit_Click()
On Error GoTo Err_cmdEdit_Click

  Me.AllowAdditions = True
  Me.AllowEdits = True
  cboFind.Enabled = False
  'Me.frmCustomer.Enabled = True
'  If Me!frmCustomer.Form.Recordset.RecordCount > 0 Then
'    Me!frmCustomer.Form.cmdEdit.Enabled = True
'  Else
'    Me!frmCustomer.Form.cmdEdit.Enabled = False
'  End If
cboSubAccount.Enabled = False

  accOpBal.Locked = False
  cboSubAccount.Locked = False
  accNo.Locked = False
  accName1.Locked = False
  accStatus.Locked = False
  accIsVat.Locked = False
  accVatable.Locked = False
  accSign.Locked = False
  accType.Locked = False
  accOpenBalance.Locked = False
  accCur.Locked = False
  accOpenBalCur.Locked = False
  accOpenBalance.Enabled = True
  accNo.SetFocus
  cmdEdit.Enabled = False
  cmdExit.Enabled = False
  cmdAdd.Enabled = True
  cmdUndo.Enabled = True
  cmdSave.Enabled = True
  cmdChart.Enabled = False
  
Exit_cmdEdit_Click:
    Exit Sub
Err_cmdEdit_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdEdit_Click
End Sub

Private Sub cmdFirst_Click()
On Error GoTo Err_cmdFirst_Click
    
If Me.AllowEdits = True Then
    Call cmdSave_Click
    'DoCmd.RunCommand acCmdSaveRecord
End If
If HasChanged Then
    If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
        Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
    End If
    
    If fsClosedYear Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
        Else
            If IsNull(accNoCopy) Then
                accNoCopy = accNo
            End If
            sUpdateAccount (accNoCopy)
        End If
    End If
End If
    
    DoCmd.GoToRecord , , acFirst
    lblPrvRef.Visible = False
'    If Me!frmCustomer.Form.Recordset.RecordCount > 0 Then
'        Me!frmCustomer.Form.cmdEdit.Enabled = True
'    Else
'        Me!frmCustomer.Form.cmdEdit.Enabled = False
'    End If
Exit_cmdFirst_Click:
    Exit Sub
Err_cmdFirst_Click:
    If Err.Number = 13 Then GoTo Exit_cmdFirst_Click
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdFirst_Click
End Sub
Private Sub cmdLast_Click()
On Error GoTo Err_cmdLast_Click
    
If Me.AllowEdits = True Then
    Call cmdSave_Click
    'DoCmd.RunCommand acCmdSaveRecord
End If
If HasChanged Then
    If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
        Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
    End If
    
    If fsClosedYear Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
        Else
            If IsNull(accNoCopy) Then
                accNoCopy = accNo
            End If
            sUpdateAccount (accNoCopy)
        End If
    End If
End If
    
    DoCmd.GoToRecord , , acLast
    lblPrvRef.Visible = False
    
'    If Me!frmCustomer.Form.Recordset.RecordCount > 0 Then
'        Me!frmCustomer.Form.cmdEdit.Enabled = True
'    Else
'        Me!frmCustomer.Form.cmdEdit.Enabled = False
'    End If
    
Exit_cmdLast_Click:
    Exit Sub
Err_cmdLast_Click:
    If Err.Number = 13 Then GoTo Exit_cmdLast_Click
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdLast_Click
End Sub
Private Sub cmdNext_Click()
On Error GoTo Err_cmdNext_Click
    
If Me.AllowEdits = True Then
    Call cmdSave_Click
    'DoCmd.RunCommand acCmdSaveRecord
End If
If HasChanged Then
    If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
        Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
    End If

    If fsClosedYear Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
            HasChanged = False
        Else
            If IsNull(accNoCopy) Then
                accNoCopy = accNo
            End If
            sUpdateAccount (accNoCopy)
        End If
    End If
End If

    If Me.Recordset.RecordCount <> Me.Recordset.AbsolutePosition + 1 Then
        DoCmd.GoToRecord , , acNext
    End If
    lblPrvRef.Visible = False
'    If Me!frmCustomer.Form.Recordset.RecordCount > 0 Then
'        Me!frmCustomer.Form.cmdEdit.Enabled = True
'    Else
'        Me!frmCustomer.Form.cmdEdit.Enabled = False
'    End If
    
Exit_cmdNext_Click:
    Exit Sub
Err_cmdNext_Click:
    Select Case Err.Number
    Case 13
    Case 2105
        Beep
    Case Else
        MsgBox FnIsErr(Err.Number), vbExclamation
    End Select
    Resume Exit_cmdNext_Click
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo Err_cmdPrevious_Click
    
If Me.AllowEdits = True Then
    Call cmdSave_Click
    'DoCmd.RunCommand acCmdSaveRecord
End If
If HasChanged Then
    If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
        Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
    End If

    If fsClosedYear Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
        Else
            If IsNull(accNoCopy) Then
                accNoCopy = accNo
            End If
            sUpdateAccount (accNoCopy)
        End If
    End If
End If
    
    lblPrvRef.Visible = False
    DoCmd.GoToRecord , , acPrevious
'    If Me!frmCustomer.Form.Recordset.RecordCount > 0 Then
'        Me!frmCustomer.Form.cmdEdit.Enabled = True
'    Else
'        Me!frmCustomer.Form.cmdEdit.Enabled = False
'    End If
Exit_cmdPrevious_Click:
    Exit Sub
Err_cmdPrevious_Click:
    Select Case Err.Number
    Case 13
    Case 2105
        Beep
    Case Else
        MsgBox FnIsErr(Err.Number), vbExclamation
    End Select
    Resume Exit_cmdPrevious_Click
End Sub
Private Sub cmdAdd_Click()
On Error GoTo Err_cmdAdd_Click
 
'If Me.NewRecord = False Then
    'DoCmd.RunCommand acCmdSaveRecord
    If HasChanged Then
        If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
            Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
        End If

        If fsClosedYear Then
            If fsNextYearHasChanged Then
                MsgBox FnIsErr(7008), vbExclamation
            Else
                If IsNull(accNoCopy) Then
                    accNoCopy = accNo
                End If
                sUpdateAccount (accNoCopy)
            End If
        End If
    End If
'End If

  Me.AllowAdditions = True
'  Me.frmCustomer.Enabled = True
  DoCmd.GoToRecord , , acNewRec
  lblPrvRef.Visible = False
  accOpBal.Locked = False
  cboSubAccount.Locked = False
  accNo.Locked = False
  accName1.Locked = False
  accStatus.Locked = False
  accIsVat.Locked = False
  accVatable.Locked = False
  accSign.Locked = False
  accSign.Value = "C"
  accType.Locked = False
  accOpenBalance.Locked = False
  accCur.Locked = False
  accOpenBalCur.Locked = False
  accOpenBalance.Enabled = True
  accNo.SetFocus
  cboSubAccount.Enabled = True
   cboFind.Enabled = False
  cmdEdit.Enabled = False
  cmdExit.Enabled = True
  cmdAdd.Enabled = True
  cmdUndo.Enabled = True
  cmdSave.Enabled = True
  cmdChart.Enabled = False
  
  pgeAcc.SetFocus
  frmRef.SetFocus
  Me.txtLev1 = Empty
  Me.txtLev2 = Empty
  Me.txtLev3 = Empty
  Me.txtLev4 = Empty
'  Me!frmCustomer.Form.cmdEdit.Enabled = False
  
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click
End Sub
Private Sub cmdDel_Click()
Dim rs As Recordset
Dim strSQL As String
Dim intaccID As Integer

On Error GoTo Err_cmdDel_Click
'    DoCmd.SetWarnings False
'    If MsgBox("Delete Record? YES or NO?", vbYesNo, "Warning!") = vbYes Then
'        DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
'        DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
'        frmRef.SetFocus
'    Else
'        Exit Sub
'    End If
'    DoCmd.SetWarnings True

If (Me.NewRecord = False And IsNull(accID.Value) = False) Then
    strSQL = "SELECT Count(accID) FROM tblTransactionSub WHERE accID = " & accID
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.Fields(0) > 0 Then
        MsgBox FnIsErr(7007), vbExclamation
        Exit Sub
    End If
Else
    'empty fields
    DoCmd.GoToRecord , , acPrevious
    Me.Refresh
    Exit Sub
End If

lblPrvRef.Visible = False
     
    If MsgBox("Delete Record? YES or NO?", vbYesNo, "Warning!") = vbYes Then
        If (IsNull(accID.Value) = False) Then
            Me.accSign = "C"
            Me.accType = 1
            intaccID = accID.Value
            DoCmd.GoToRecord , , acPrevious
            DoCmd.SetWarnings False
            strSQL = "DELETE from tblAccount WHERE accID=" & intaccID
            FnLog (strSQL)
            DoCmd.RunSQL (strSQL)
            'changes the account if exists in customer
            If DLookup("ctrClosed", "tblControl") = False Then
                strSQL = "UPDATE appCustomer SET accID = 0 WHERE accID = " & intaccID
                On Error Resume Next
                DoCmd.RunSQL (strSQL)
                FnLog ("********OnErrorNoExecute********* " & strSQL)
            End If
            DoCmd.SetWarnings True
        Else
            Me.Undo
            'Call cmdUndo_Click
            Exit Sub
        End If
        'refreshes and finds previous record
        intaccID = accID.Value
        'Me.Refresh
        Me.Requery
        Set rs = Me.Recordset.Clone
        rs.FindFirst "[accID] = " & intaccID
        If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    Else
    End If
    
'    If Me!frmCustomer.Form.Recordset.RecordCount > 0 Then
'        Me!frmCustomer.Form.cmdEdit.Enabled = True
'    Else
'        Me!frmCustomer.Form.cmdEdit.Enabled = False
'    End If
Exit_cmdDel_Click:
    Exit Sub
Err_cmdDel_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdDel_Click
End Sub
Private Sub cmdExit_Click()

If HasChanged Then
    If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
        Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
    End If
    If fsClosedYear Then
        If fsNextYearHasChanged Then
            MsgBox FnIsErr(7008), vbExclamation
        Else
            If IsNull(accNoCopy) Then
                accNoCopy = accNo
            End If
            sUpdateAccount (accNoCopy)
        End If
    End If
End If

If IsNoData(Me.OpenArgs) Then
    DoCmd.OpenForm "frmMain"
    DoCmd.Close acForm, Me.name
Else
    Select Case Me.OpenArgs
        Case "frmTransactionSub"
            Forms.frmTransaction.frmSub.Controls("accID1").Value = accID
            DoCmd.Close acForm, "frmAccount"
            Screen.ActiveForm.Refresh
        Case "frmControl"
            Forms.frmControl.Controls("ctrLYPLA").Value = accID
            DoCmd.Close acForm, "frmAccount"
            Screen.ActiveForm.Refresh
        Case "frmAccCust"
            Forms!frmTransControl!frmAccCust.Form.cboCustID.Requery
            Forms!frmTransControl!frmAccCust.Form.accID = accID
            DoCmd.Close acForm, "frmAccount"
            'Screen.ActiveForm.Refresh
        Case Else
    
    
    End Select
End If


End Sub

Private Sub cmdSave_Click()
  On Error GoTo Err_End
'    If Me.frmCustomer.Form.cmdDel.Enabled Then
'        Call Me.frmCustomer.Form.cmdSave_Click
'    End If

    If Me.NewRecord _
    And (IsNull(coaID) = True Or IsEmpty(coaID) = True) _
    And (IsNull(accNo) = True Or IsEmpty(accNo) = True) Then
        DoCmd.GoToRecord , , acPrevious
        Me.cboSubAccount.Requery
    Else
        'DoCmd.RunCommand acCmdSave
        'DoCmd.RunCommand acCmdSaveRecord
    End If
    
    If ((IsNumeric(cboSubAccount)) And (cboSubAccount > 0)) Then
        Call ChartAccType(cboSubAccount.Column(5), cboSubAccount.Column(6), cboSubAccount.Column(7), cboSubAccount.Column(8))
    End If
    
    Me.cmdNext.SetFocus
    'Me.frmCustomer.Enabled = False
    
    'cmdDel.Enabled = False
    cmdExit.Enabled = True
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdUndo.Enabled = False
    cmdChart.Enabled = True
    cmdSave.Enabled = False
    
    'check the customers data - if there are no data, this is new record and custID exists then delete it
    If Len(Nz(Me.custName, "")) = 0 And Me.lblPos.Caption = "Rec. New" And IsNull(Me.custID) = False Then
        Dim strSQL As String
        strSQL = "DELETE FROM appCustomer WHERE custID = " & Me.custID
        DoCmd.SetWarnings False
        DoCmd.RunSQL (strSQL)
        DoCmd.SetWarnings True
        
        Dim lngaccID As Long
        Dim byttabAcc As Byte
        lngaccID = Me.accID
        byttabAcc = Me.tabAcc
        Me.Requery
        Me.Refresh
        Me.Recordset.FindFirst ("accID = " & lngaccID)
        Me.tabAcc = byttabAcc
    End If
    
   ' DoCmd.RunCommand acCmdSaveRecord
    
    cboSubAccount.Enabled = False
     cboFind.Enabled = True
    Me.AllowAdditions = False
    Me.AllowEdits = False
    
    
    Exit Sub
Err_End:
    If Err.Number <> 0 Then
        MsgBox FnIsErr(Err.Number), vbExclamation
    Else
    End If
End Sub

Private Sub cmdUndo_Click()
On Error GoTo Err_End
    DoCmd.RunCommand acCmdUndo
    cmdSave.Enabled = False
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdExit.Enabled = True
    cmdDel.Enabled = False
    Me.AllowAdditions = False
    Me.AllowEdits = False
    cmdChart.Enabled = True
    'Me.frmCustomer.Enabled = False
    accName1.SetFocus
    cmdUndo.Enabled = False
Exit Sub
Err_End:
    If Err.Number <> 0 Then
        MsgBox FnIsErr(Err.Number), vbExclamation
    Else
    End If
End Sub

Private Sub custBankAcc_Click()
    If Me.custBankAcc = True Then
        Me.custBankAccNo.Visible = True
    Else
        Me.custBankAccNo.Visible = False
    End If
End Sub



Private Sub Form_AfterUpdate()
'Me.accNoCopy = Me.accNo
End Sub

Private Sub Form_Current()
' Purpose: Show current record position
If Me.NewRecord Then
    Me!lblPos.Caption = "Rec. New"
Else
    Me.RecordsetClone.Bookmark = Me.Bookmark
    Me!lblPos.Caption = "Rec. " & CStr(Me.RecordsetClone.AbsolutePosition + 1) _
                        & " of " & CStr(Me.RecordsetClone.RecordCount)
End If
Me.Caption = FnIme
cboFind.Requery

If Me.NewRecord = False Then
    Dim rs As Recordset
    Dim strSQL As String
    Dim arrPer As Variant
    
    'sets to show all refereneces in chart of account
    strSQL = "SELECT tblChart.coaID, tblChart.coaName, tblChart.coaNo, tblChart.coaIDg, tblChart.lvlID, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart.coaRef5 " & _
           "FROM tblChart " & _
           "WHERE tblChart.lvlID = 5 " & _
           "ORDER BY tblChart.coaName, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart.coaRef5; "
    Me.cboSubAccount.RowSource = strSQL
    Me.cboSubAccount.Requery
    
    If IsNull(cboSubAccount.Column(5)) = False Then
        strSQL = "SELECT tblChart.coaID, tblChart.coaNo, tblChart.coaName, Right(Chr(48) & CStr(tblChart.coaRef1), 2) AS coaRef1, Right(Chr(48) & CStr(tblChart.coaRef2), 2) AS coaRef2, Right(Chr(48) & CStr(tblChart.coaRef3), 2) AS coaRef3, Right(Chr(48) & CStr(tblChart.coaRef4), 2) As coaRef4, tblChart.lvlID FROM tblChart WHERE (tblChart.coaRef1 = " & cboSubAccount.Column(5) & " And tblChart.lvlID = 1) OR (tblChart.coaRef1 & tblChart.coaRef2 =  " & cboSubAccount.Column(5) & cboSubAccount.Column(6) & " And tblChart.lvlID = 2) OR (tblChart.coaRef1 & tblChart.coaRef2& tblChart.coaRef3 =  " & cboSubAccount.Column(5) & cboSubAccount.Column(6) & cboSubAccount.Column(7) & " And tblChart.lvlID = 3) OR (tblChart.coaRef1 & tblChart.coaRef2& tblChart.coaRef3& tblChart.coaRef4 =  " & cboSubAccount.Column(5) & cboSubAccount.Column(6) & cboSubAccount.Column(7) & cboSubAccount.Column(8) & " And tblChart.lvlID = 4) ORDER BY tblChart.lvlID"

        Set rs = CurrentDb.OpenRecordset(strSQL)

        txtLev1 = ""
        txtLev2 = ""
        txtLev3 = ""
        txtLev4 = ""

        Do While (Not rs.EOF)
            Select Case rs.Fields("lvlID")
            Case 1
                txtLev1 = Right("0" & rs.Fields("coaRef1"), 2) & "  " & rs.Fields("coaName")
            Case 2
                txtLev2 = Right("0" & rs.Fields("coaRef2"), 2) & "  " & rs.Fields("coaName")
            Case 3
                txtLev3 = Right("0" & rs.Fields("coaRef3"), 2) & "  " & rs.Fields("coaName")
            Case 4
                txtLev4 = Right("0" & rs.Fields("coaRef4"), 2) & "  " & rs.Fields("coaName")
            End Select
            rs.MoveNext
        Loop
            rs.Close
    End If

    If IsEmpty(accID) = False Then
        strSQL = "SELECT tblPeriod.perID, Sum(IIf([trnsSign]='C',-1*[trnsAmount],[trnsAmount])) AS Bal, tblAccount.accID " & _
                 "FROM tblAccount INNER JOIN (tblTransactionSub INNER JOIN (tblTransaction INNER JOIN (tblPeriodSel INNER JOIN tblPeriod ON tblPeriodSel.perID = tblPeriod.perID) ON tblTransaction.persID = tblPeriodSel.persID) ON tblTransactionSub.trnID = tblTransaction.trnID) ON tblAccount.accID = tblTransactionSub.accID " & _
                 "WHERE tblAccount.accID = " & Me.accID.Value & _
                " AND tblTransaction.trnYear in (SELECT yearID FROM tblControl) " & _
                " GROUP BY tblPeriod.perID, tblAccount.accID " & _
                "ORDER BY tblPeriod.perID;"
  
        arrPer = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        Set rs = CurrentDb.OpenRecordset(strSQL)

        Do While (Not rs.EOF)
            arrPer(rs.Fields("perID") - 1) = rs.Fields("Bal")
            rs.MoveNext
        Loop
            rs.Close
        arrPer(0) = arrPer(0) + Me.accOpenBalance

        For i = 1 To 11
            arrPer(i) = arrPer(i) + arrPer(i - 1)
        Next i

        Me.accPer1.Value = arrPer(0)
        Me.accPer2.Value = arrPer(1)
        Me.accPer3.Value = arrPer(2)
        Me.accPer4.Value = arrPer(3)
        Me.accPer5.Value = arrPer(4)
        Me.accPer6.Value = arrPer(5)
        Me.accPer7.Value = arrPer(6)
        Me.accPer8.Value = arrPer(7)
        Me.accPer9.Value = arrPer(8)
        Me.accPer10.Value = arrPer(9)
        Me.accPer11.Value = arrPer(10)
        Me.accPer12.Value = arrPer(11)
    
        OldValue = cboSubAccount.Column(1) & " " _
                 & Right("0" & CStr(cboSubAccount.Column(5)), 2) & "-" _
                 & Right("0" & CStr(cboSubAccount.Column(6)), 2) & "-" _
                 & Right("0" & CStr(cboSubAccount.Column(7)), 2) & "-" _
                 & Right("0" & CStr(cboSubAccount.Column(8)), 2) & "-" _
                 & Right("00" & CStr(cboSubAccount.Column(9)), 3)
        OldId = cboSubAccount.Column(0)
        lblPrvRef.Visible = False
    End If
        
Else
    Me.accPer1.Value = Me.accOpenBalance
    Me.accPer2.Value = Me.accOpenBalance
    Me.accPer3.Value = Me.accOpenBalance
    Me.accPer4.Value = Me.accOpenBalance
    Me.accPer5.Value = Me.accOpenBalance
    Me.accPer6.Value = Me.accOpenBalance
    Me.accPer7.Value = Me.accOpenBalance
    Me.accPer8.Value = Me.accOpenBalance
    Me.accPer9.Value = Me.accOpenBalance
    Me.accPer10.Value = Me.accOpenBalance
    Me.accPer11.Value = Me.accOpenBalance
    Me.accPer12.Value = Me.accOpenBalance
    
    'sets to show only available refereneces in chart of account
    strSQL = "SELECT tblChart.coaID, tblChart.coaName, tblChart.coaNo, tblChart.coaIDg, tblChart.lvlID, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart.coaRef5 " & _
             "FROM tblChart " & _
             "WHERE (((tblChart.coaID) Not In (SELECT tblAccount.caoID FROM tblAccount)) AND ((tblChart.lvlID)=5)) " & _
             "ORDER BY tblChart.coaName, tblChart.coaRef1, tblChart.coaRef2, tblChart.coaRef3, tblChart.coaRef4, tblChart.coaRef5; "
    Me.cboSubAccount.RowSource = strSQL
    Me.cboSubAccount.Requery
End If


End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim strMsg As String
        strMsg = FnIsErr(DataErr)
        Response = acDataErrContinue
        If IsNoData(strMsg) = False Then
            MsgBox strMsg, vbExclamation, "Error!"
        Else
        End If
End Sub


Private Sub Form_Open(Cancel As Integer)
    Select Case Me.RecordsetClone.RecordCount
        Case 0
                Me!lblPos.Caption = "Rec. New"
        Case Else
            If Me.NewRecord Then
                Me!lblPos.Caption = "Rec. New"
            Else
                Me.RecordsetClone.Bookmark = Me.Bookmark
                Me!lblPos.Caption = "Rec. " & CStr(Me.RecordsetClone.AbsolutePosition + 1) _
                                    & " of " & CStr(Me.RecordsetClone.RecordCount)
            End If
    End Select
    
    Me.AllowAdditions = True
    
    Select Case Me.OpenArgs
        Case "frmTransactionSub":
            'Me.Recordset.FindFirst "[accID] = " & Forms.frmTransaction.frmSub.Controls("accID1").Value
            Me.cboFind.Value = Forms.frmTransaction.frmSub.Controls("accID1").Value
            cboFind_AfterUpdate
        Case "frmAccCust":
            Call cmdAdd_Click
    End Select
    
    If ((fsTableExists("appCustomer") And (fsTableExists("appCustomer")))) Then
        accEasyDoc.Visible = True
    Else
        accEasyDoc.Visible = False
    End If

    Me.cmdExit.SetFocus
End Sub

Public Sub GetReference(NewcoaID As Long)
   Me.caoID = NewcoaID
   cboSubAccount = NewcoaID
   accNo.SetFocus
   cboSubAccount_Change
   cboSubAccount.Requery
   cboSubAccount_Click
   'Me.Refresh
End Sub

Private Sub grpFind_AfterUpdate()
Select Case grpFind
Case 1
      'cboFind.InputMask = ""
      cboFind.ColumnWidths = "0cm;6cm;2cm"
      cboFind.RowSource = "SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo " & _
            " FROM tblAccount ORDER BY tblAccount.accName;"
    Case Else
    cboFind.ColumnWidths = "0cm;2cm;6cm"
      cboFind.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
            " FROM tblAccount ORDER BY tblAccount.accNo;"
    
End Select
End Sub


Private Sub tabAcc_Change()
    If Me.NewRecord And IsNull(Me.accID) = False And tabAcc = 2 Then
        If Me.AllowEdits = False And Me.cmdSave.Enabled = True Then
            Me.AllowEdits = True
        End If
        DoCmd.RunCommand acCmdSaveRecord
        
'        Dim lngaccID As Long
'        lngaccID = Me.accID
'        Me.Requery
'        Me.Refresh
'        Me.Recordset.FindFirst ("accID = " & lngaccID)
'        Me.tabAcc = 2
'
'        Dim strSQL As String
'        strSQL = "INSERT INTO appCustomer(accID) VALUES (" & Me.accID & ")"
'        DoCmd.SetWarnings False
'        DoCmd.RunSQL (strSQL)
'        DoCmd.SetWarnings True
    Else
    End If

End Sub
