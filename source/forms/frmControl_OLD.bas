Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =9615
    DatasheetFontHeight =10
    ItemSuffix =274
    Left =4740
    Top =2115
    Right =14355
    Bottom =8805
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5f2f758f1b99e340
    End
    RecordSource ="SELECT tblControl.*, sysMyCo.smycoName, sysMyCo.smycoPath, sysMyCo.smycoYear FRO"
        "M tblControl INNER JOIN sysMyCo ON tblControl.ctrDBID=sysMyCo.smycoID; "
    Caption ="Control File"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa4050000a2050000a2050000a205000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =755
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =465
                    Top =180
                    Width =7755
                    Height =480
                    Name ="lblComent"
                    Caption ="System Parameters and Company Informations"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5460
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Width =9615
                    Height =5460
                    Name ="tabAcc"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =377
                            Width =9345
                            Height =4948
                            Name ="pgeGen"
                            Caption ="Account Details"
                            Begin
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =4943
                                    Top =377
                                    Width =180
                                    Height =255
                                    Name ="ctrID"
                                    ControlSource ="ctrID"
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2105
                                    Top =721
                                    Width =3330
                                    Height =255
                                    TabIndex =1
                                    Name ="ctrName"
                                    ControlSource ="ctrName"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =665
                                            Top =721
                                            Width =1380
                                            Height =255
                                            Name ="bnkaABA_Label"
                                            Caption ="Company Name"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2105
                                    Top =1051
                                    Width =3330
                                    Height =255
                                    TabIndex =2
                                    Name ="ctrDate"
                                    ControlSource ="ctrDate"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =665
                                            Top =1051
                                            Width =1380
                                            Height =255
                                            Name ="bnkaBLZ_Label"
                                            Caption ="Date"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2105
                                    Top =1359
                                    Width =3330
                                    Height =255
                                    TabIndex =3
                                    Name ="ctrPageCounter"
                                    ControlSource ="ctrPageCounter"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =665
                                            Top =1359
                                            Width =1380
                                            Height =255
                                            Name ="Label84"
                                            Caption ="Account Name"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =1134
                                    Left =2767
                                    Top =2445
                                    Width =1740
                                    Height =270
                                    TabIndex =4
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="persID"
                                    ControlSource ="persID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="1;Profit & Lose;2;Balance Sheet"
                                    ColumnWidths ="0;1134"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =943
                                            Top =2449
                                            Width =1080
                                            Height =255
                                            Name ="Label35"
                                            Caption ="Current Period"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =2112
                                    Top =3304
                                    Width =3360
                                    Height =270
                                    TabIndex =5
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="ctrLYPLA"
                                    ControlSource ="ctrLYPLA"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=1)) ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =658
                                            Top =3305
                                            Width =1395
                                            Height =255
                                            Name ="lblLYPLA"
                                            Caption ="Last Profit &&  Loose"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    Left =2113
                                    Top =2464
                                    Width =507
                                    Height =270
                                    TabIndex =6
                                    Name ="TXTpersID"
                                    ControlSource ="persID"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2862
                                    Top =2054
                                    Width =1296
                                    Height =351
                                    TabIndex =7
                                    Name ="btnCloseYear"
                                    Caption ="Close Year >>"
                                    OnClick ="[Event Procedure]"
                                    ObjectPalette = Begin
                                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                                        0xffffff0000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2106
                                    Top =2114
                                    Width =705
                                    TabIndex =8
                                    Name ="yearID"
                                    ControlSource ="yearID"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1421
                                            Top =2138
                                            Width =615
                                            Height =240
                                            Name ="Label97"
                                            Caption ="Year"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =4748
                                    Top =1923
                                    Width =510
                                    TabIndex =9
                                    Name ="smycoPath"
                                    ControlSource ="smycoPath"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =2112
                                    Top =3645
                                    Width =3360
                                    Height =270
                                    TabIndex =10
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="ctrControlAcc"
                                    ControlSource ="ctrControlAcc"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount OR"
                                        "DER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =525
                                            Top =3645
                                            Width =1530
                                            Height =255
                                            Name ="lblControlAcc"
                                            Caption ="Control Account"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2100
                                    Top =2790
                                    Width =2001
                                    Height =351
                                    TabIndex =11
                                    Name ="cmdUpdtBal"
                                    Caption ="Update Balances"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4230
                                    Top =2790
                                    Width =2001
                                    Height =351
                                    TabIndex =12
                                    Name ="cmdUpdChart"
                                    Caption ="Update Chart of Accounts"
                                    OnClick ="[Event Procedure]"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =405
                            Width =9345
                            Height =4920
                            Name ="Lock Pages"
                            EventProcPrefix ="Lock_Pages"
                            Caption ="Lock Pages"
                            Begin
                                Begin OptionGroup
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    Left =165
                                    Top =450
                                    Width =8233
                                    Height =4332
                                    Name ="frmLockPages"
                                    DefaultValue ="1"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =401
                                            Top =713
                                            OptionValue =1
                                            Name ="optRange"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =684
                                                    Top =687
                                                    Width =810
                                                    Height =240
                                                    Name ="Label123"
                                                    Caption ="Ranges"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =401
                                            Top =1215
                                            OptionValue =2
                                            Name ="optSelective"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =692
                                                    Top =1187
                                                    Width =1530
                                                    Height =240
                                                    Name ="Label125"
                                                    Caption ="Select pages to lock"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =270
                                    Top =570
                                    Width =8005
                                    Height =476
                                    Name ="Box133"
                                End
                                Begin Subform
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =285
                                    Top =1470
                                    Width =6626
                                    Height =3141
                                    TabIndex =1
                                    Name ="frmTransactionLock"
                                    SourceObject ="Form.frmTransactionLock"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3054
                                    Top =697
                                    Width =825
                                    TabIndex =2
                                    Name ="txtFromPage"
                                    Format ="General Number"
                                    ValidationRule =">=0"
                                    ValidationText ="The value must be equal or above 0"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            Left =1590
                                            Top =697
                                            Width =1320
                                            Height =240
                                            Name ="Label137"
                                            Caption ="Page ranges from"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =4421
                                    Top =685
                                    Width =780
                                    TabIndex =3
                                    Name ="txtToPage"
                                    Format ="General Number"
                                    ValidationRule =">=0"
                                    ValidationText ="The value must be equal or above 0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =4030
                                            Top =704
                                            Width =275
                                            Height =247
                                            Name ="Label143"
                                            Caption ="to"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =6731
                                    Top =690
                                    Width =720
                                    TabIndex =4
                                    Name ="txtUpToPrd"
                                    Format ="General Number"
                                    ValidationRule =">=0"
                                    ValidationText ="The value must be equal or above 0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5671
                                            Top =690
                                            Width =945
                                            Height =240
                                            Name ="Label141"
                                            Caption ="Up to period"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7515
                                    Top =638
                                    Width =661
                                    Height =330
                                    TabIndex =5
                                    Name ="cmdStartRge"
                                    Caption ="Start"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7185
                                    Top =4275
                                    Width =969
                                    Height =331
                                    TabIndex =6
                                    Name ="cmdCommit"
                                    Caption ="Commit"
                                    OnClick ="[Event Procedure]"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =405
                            Width =9345
                            Height =4920
                            Name ="pgePeriod"
                            Caption ="Current Balance"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =712
                                    Top =508
                                    Width =4995
                                    Height =4590
                                    Name ="Reference File"
                                    SourceObject ="Form.frmPeriod"
                                    EventProcPrefix ="Reference_File"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5975
                                    Top =3757
                                    Width =2265
                                    Height =900
                                    Name ="Label106"
                                    Caption ="NOTICE:\015\012Any change in \"Period Section will direclt take effect in calcul"
                                        "ation.\015\012"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =405
                            Width =9345
                            Height =4920
                            Name ="pgeInterface"
                            Caption ="Interface"
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =512
                                    Width =3360
                                    Height =270
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboInvCR"
                                    ControlSource ="ctrInvCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE tblAccount.accType=1 ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =512
                                            Width =1530
                                            Height =255
                                            Name ="lblInvCR"
                                            Caption ="CR Sales A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =512
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblInv"
                                    Caption ="INVOICES && CASH SALES"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =465
                                    Width =7678
                                    Height =378
                                    Name ="Box193"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =1427
                                    Width =3360
                                    Height =270
                                    TabIndex =1
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboCRN_DR"
                                    ControlSource ="ctrCRN_DR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE tblAccount.accType=1 ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =1427
                                            Width =1530
                                            Height =255
                                            Name ="lblCRN_DR"
                                            Caption ="DR Sales A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =1427
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblCRN"
                                    Caption ="CREDIT NOTES"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =1380
                                    Width =7678
                                    Height =378
                                    Name ="Box205"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =1885
                                    Width =3360
                                    Height =270
                                    TabIndex =2
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboDRN_CR"
                                    ControlSource ="ctrDRN_CR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE tblAccount.accType=1 ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =1885
                                            Width =1530
                                            Height =255
                                            Name ="lblDRN_CR"
                                            Caption ="CR Sales A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =1885
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblDRN"
                                    Caption ="DEBIT NOTES"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =1838
                                    Width =7678
                                    Height =378
                                    Name ="Box209"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =2343
                                    Width =3360
                                    Height =270
                                    TabIndex =3
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboRcptDR"
                                    ControlSource ="ctrRcptDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE tblAccount.accType=1 ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =2343
                                            Width =1530
                                            Height =255
                                            Name ="lblRcptDR"
                                            Caption ="DR Cost A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =2343
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblRcpt"
                                    Caption ="RECEIPTS"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =2296
                                    Width =7678
                                    Height =378
                                    Name ="Box213"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =2801
                                    Width =3360
                                    Height =270
                                    TabIndex =4
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboExpDR"
                                    ControlSource ="ctrExpDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE tblAccount.accType=1 ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =2801
                                            Width =1530
                                            Height =255
                                            Name ="lblExpDR"
                                            Caption ="DR Purchase A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =2801
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblExp"
                                    Caption ="EXPENSES"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =2754
                                    Width =7678
                                    Height =378
                                    Name ="Box217"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =3259
                                    Width =3360
                                    Height =270
                                    TabIndex =5
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboPaymtCR"
                                    ControlSource ="ctrPaymtCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2638
                                            Top =3255
                                            Width =1635
                                            Height =255
                                            Name ="lblPaymtCR"
                                            Caption ="CR Cash or Bank A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =3259
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblPaymt"
                                    Caption ="PAYMENTS"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =3212
                                    Width =7678
                                    Height =378
                                    Name ="Box221"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =3717
                                    Width =3360
                                    Height =270
                                    TabIndex =6
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboDeptsCR"
                                    ControlSource ="ctrDeptsCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =3717
                                            Width =1530
                                            Height =255
                                            Name ="lblDeptsCR"
                                            Caption ="CR Cash A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =3717
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblDepts"
                                    Caption ="DEPOSITS"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =3670
                                    Width =7678
                                    Height =378
                                    Name ="Box225"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =4175
                                    Width =3360
                                    Height =270
                                    TabIndex =7
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboCshWithdrwDR"
                                    ControlSource ="ctrCshWithdrwDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =4175
                                            Width =1530
                                            Height =255
                                            Name ="lblCshWithdrwDR"
                                            Caption ="DR Cash A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =4175
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblWDR"
                                    Caption ="WITHDRAWALS"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =4128
                                    Width =7678
                                    Height =378
                                    Name ="Box229"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =240
                                    Top =4752
                                    TabIndex =8
                                    Name ="chkDocAddYear"
                                    ControlSource ="ctrDocAddYear"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =475
                                            Top =4717
                                            Width =2535
                                            Height =240
                                            Name ="Label232"
                                            Caption ="Include Year in document numbers"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4359
                                    Top =977
                                    Width =3360
                                    Height =270
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                                    Name ="cboCSH_DR"
                                    ControlSource ="ctrCSH_DR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE tblAccount.accType=1 ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2740
                                            Top =977
                                            Width =1530
                                            Height =255
                                            Name ="Label271"
                                            Caption ="DR Sales A/C"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =253
                                    Top =977
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="Label272"
                                    Caption ="CASH SALES"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =165
                                    Top =930
                                    Width =7678
                                    Height =378
                                    Name ="Box273"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =45
                            Top =360
                            Width =9435
                            Height =4965
                            Name ="pgeCust"
                            Caption ="Customers"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =45
                                    Top =360
                                    Width =9435
                                    Height =4800
                                    Name ="frmAccCust"
                                    SourceObject ="Form.frmAccCust"
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =600
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =70
                    Left =210
                    Top =188
                    Width =366
                    Name ="cmdFirst"
                    Caption ="&F"
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
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="First Record (Alt+F)"
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =76
                    Left =1470
                    Top =188
                    Width =366
                    TabIndex =1
                    Name ="cmdLast"
                    Caption ="&L"
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
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Last Record (Alt+L)"
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =1050
                    Top =188
                    Width =366
                    TabIndex =2
                    Name ="cmdNext"
                    Caption ="&N"
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
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Next Record (Alt+N)"
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    Left =630
                    Top =188
                    Width =366
                    TabIndex =3
                    Name ="cmdPrevious"
                    Caption ="&P"
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
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Previous Record (Alt+P)"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =4751
                    Top =180
                    Width =1080
                    TabIndex =4
                    Name ="cmdAdd"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =68
                    Left =5951
                    Top =180
                    Width =1080
                    TabIndex =5
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =7151
                    Top =180
                    Width =1080
                    TabIndex =6
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form"
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

Private Sub btnCloseYear_Click()
Dim strNewDBPath As String
Dim Db As Database
Dim rs As DAO.Recordset
Dim strSQL As String, dblBalance As Double

Const cERR_USERCANCEL = vbObjectError + 1000

strNewDBPath = CreateNewMDBFile

If fRefreshLinks(smycoPath, True) Then

    DoCmd.SetWarnings False
    
    'Updates local control table
    strSQL = "INSERT INTO sysMyCo (smycoName, smycoPath, smycoYear) " & _
             "VALUES ('" & Me.smycoName & "', '" & strNewDBPath & "', " & Me.smycoYear + 1 & ")"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    'Sets flag to closed year
    strSQL = "UPDATE tblControl " & _
                 "SET ctrClosed = True"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    'Updates company control table
    strSQL = "SELECT smycoID FROM sysMyCo WHERE smycoPath = '" & strNewDBPath & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    rs.MoveFirst
    
    strSQL = "UPDATE nexttblControl " & _
             "SET yearID = " & Me.Recordset.Fields("yearID") + 1 & ", " & _
             "persID = 1, " & _
             "ctrDBID = " & rs.Fields("smycoID")
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
   
    'Clean transactions
    strSQL = "DELETE * FROM nexttblTransaction"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    
    strSQL = "DELETE * FROM nexttblTransactionSub"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)

    'Updates balances

    
    sUpdtBalances
    
    'Check if balances match
    
    strSQL = "SELECT Sum(a.Bal) AS currBal, b.nextBal FROM (" & _
             "SELECT (Sum(IIf([trnsSign]='C',[trnsAmount],-1*[trnsAmount]))) AS Bal " & _
             "FROM (tblAccount INNER JOIN (tblTransactionSub INNER JOIN tblTransaction ON tblTransactionSub.trnID=tblTransaction.trnID) ON tblAccount.accID=tblTransactionSub.accID) INNER JOIN nexttblAccount ON tblAccount.accID=nexttblAccount.accID " & _
             "WHERE (((tblTransaction.trnYear) In (SELECT yearID FROM tblControl))) " & _
             "UNION ALL " & _
             "SELECT  Sum(tblAccount.accOpenBalance) AS Bal " & _
             "FROM tblAccount) a, (SELECT Sum(nexttblAccount.accOpenBalance) As nextBal " & _
             "FROM nexttblAccount) b " & _
             "GROUP BY b.nextBal " & _
             "HAVING Sum(a.Bal) <> b.nextBal"
             
    Set Db = CurrentDb()
    Set rs = Db.OpenRecordset(strSQL)
    If rs.RecordCount Then
        MsgBox FnIsErr(7005), vbExclamation
    Else
        If MsgBox("New year has been succesfully created. Do you want to open it?", _
                vbQuestion + vbYesNo, "Please confirm...") = vbNo Then
            Me.cmdExit.SetFocus
            Me.btnCloseYear.Enabled = False
        Else
            If fRefreshLinks(strNewDBPath, True) Then
                DoCmd.Close acForm, "frmControl"
                DoCmd.OpenForm "frmControl"
            Else
                MsgBox FnIsErr(7006), vbExclamation
            End If
        End If
    End If
    
    rs.Close
    Db.Close
    Set rs = Nothing
    Set Db = Nothing
    DoCmd.SetWarnings True
End If
End Sub

Private Sub cboCRN_DR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboCshWithdrwDR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboDeptsCR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboDRN_CR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboExpDR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboInvCR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboPaymtCR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cboRcptDR_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub cmdCommit_Click()
Dim bolSuccess As Boolean

Me.frmTransactionLock.Requery
bolSuccess = fsSequence(0, 0)

End Sub

Private Sub cmdFirst_Click()
On Error GoTo Err_cmdFirst_Click
    DoCmd.GoToRecord , , acFirst
Exit_cmdFirst_Click:
    Exit Sub
Err_cmdFirst_Click:
    If Err.Number = 13 Then GoTo Exit_cmdFirst_Click
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdFirst_Click
End Sub
Private Sub cmdLast_Click()
On Error GoTo Err_cmdLast_Click
    DoCmd.GoToRecord , , acLast
Exit_cmdLast_Click:
    Exit Sub
Err_cmdLast_Click:
    If Err.Number = 13 Then GoTo Exit_cmdLast_Click
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdLast_Click
End Sub
Private Sub cmdNext_Click()
On Error GoTo Err_cmdNext_Click
    DoCmd.GoToRecord , , acNext
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
    DoCmd.GoToRecord , , acPrevious
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
  Me.AllowAdditions = True
  DoCmd.GoToRecord , , acNewRec
  frmRef1.SetFocus
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click
End Sub
Private Sub cmdDel_Click()
On Error GoTo Err_cmdDel_Click
    DoCmd.SetWarnings False
    If MsgBox("Delete Record? YES or NO?", vbYesNo, "Warning!") = vbYes Then
        DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
        DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
        frmRef1.SetFocus
    Else
        Exit Sub
    End If
    DoCmd.SetWarnings True
Exit_cmdDel_Click:
    Exit Sub
Err_cmdDel_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdDel_Click
End Sub
Private Sub cmdExit_Click()
If IsNoData(Me.OpenArgs) Then
    DoCmd.OpenForm "frmMain"
    DoCmd.Close acForm, Me.name
Else
    Select Case Me.OpenArgs
        Case "frmTransactionSub"
            DoCmd.Close acForm, "frmAccount"
            Screen.ActiveForm.Refresh
        Case Else
    End Select
End If
End Sub

Private Sub cmdStartRge_Click()
Dim strPages As String
Dim i As Integer
Dim bolSuccess As Boolean

If IsNull(Me.txtFromPage) Then
    MsgBox ("Fill begining range for pages.")
    Exit Sub
End If

If IsNull(Me.txtToPage) Then
    MsgBox ("Fill ending range for pages.")
    Exit Sub
End If

Dim strSQL As String

strSQL = "UPDATE tblTransaction " & _
         "SET tblTransaction.trnLock = True " & _
         "WHERE tblTransaction.trnPageCounter Between " & Me.txtFromPage & " AND " & Me.txtToPage & _
         " AND tblTransaction.persID <= " & Me.txtUpToPrd
      
DoCmd.SetWarnings False
FnLog (strSQL)
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings True

bolSuccess = fsSequence(CInt(Me.txtFromPage), CInt(Me.txtToPage))
Me.frmTransactionLock.Requery
End Sub

Private Sub cmdUpdChart_Click()
    Dim rs As Recordset
    
    
    strSQL = "SELECT coaRef1, coaRef2, coaRef3, coaRef4 " & _
             "FROM tblChart " & _
             "WHERE lvlID = 5"
             
    Set rs = CurrentDb.OpenRecordset(strSQL)
    rs.MoveFirst
    Do Until rs.EOF
        Call ChartAccType(rs.Fields("coaRef1"), rs.Fields("coaRef2"), rs.Fields("coaRef3"), rs.Fields("coaRef4"))
        rs.MoveNext
    Loop
    DoCmd.SetWarnings False
    strSQL = "UPDATE tblChart SET coaAccType = 0 WHERE IsNull(coaAccType)"
    FnLog (strSQL)
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings True
End Sub

Private Sub cmdUpdtBal_Click()
If fsClosedYear Then
    If fsNextYearHasChanged Then
        MsgBox FnIsErr(7009), vbExclamation
    End If
    sUpdtBalances
End If

End Sub

Private Sub ctrControlAcc_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
End Sub

Private Sub ctrLYPLA_DblClick(Cancel As Integer)
DoCmd.OpenForm "frmAccount", , , , , , Me.Form.name
Me.ctrLYPLA.Requery
End Sub

Private Sub ctrLYPLA_Exit(Cancel As Integer)
    FnAllocAC (Nz(Me.ActiveControl, 0))
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


Private Sub Form_Load()
Me.txtUpToPrd = Me.persID
If Me.Recordset.Fields("ctrClosed") = True Then
'    Me.btnCloseYear.Enabled = False
End If
End Sub

Private Sub frmLockPages_Click()
If Me.frmLockPages = 1 Then
    Me.txtFromPage.Enabled = True
    Me.txtToPage.Enabled = True
    Me.txtUpToPrd.Enabled = True
    Me.cmdStartRge.Enabled = True
    Me.txtFromPage.SetFocus
    
    Me.frmLockPages.Requery
    Me.frmTransactionLock.Enabled = False
    Me.frmTransactionLock.Form.trnEntryDate.ForeColor = 8421504
    Me.frmTransactionLock.Form.trnInternalRef.ForeColor = 8421504
    Me.frmTransactionLock.Form.trnPageCounter.ForeColor = 8421504
    Me.frmTransactionLock.Form.persID.ForeColor = 8421504
    Me.frmTransactionLock.Form.trnYear.ForeColor = 8421504
    Me.frmTransactionLock.Form.trnEntryDate.BorderColor = 8421504
    Me.frmTransactionLock.Form.trnInternalRef.BorderColor = 8421504
    Me.frmTransactionLock.Form.trnPageCounter.BorderColor = 8421504
    Me.frmTransactionLock.Form.persID.BorderColor = 8421504
    Me.frmTransactionLock.Form.trnYear.BorderColor = 8421504
    Me.cmdCommit.Enabled = False
    If Me.frmTransactionLock.Form.Recordset.RecordCount > 0 Then
        Me.frmTransactionLock.Form.trnLock = 0
    End If
Else
    Me.txtFromPage.Enabled = False
    Me.txtToPage.Enabled = False
    Me.txtUpToPrd.Enabled = False
    Me.cmdStartRge.Enabled = False
    
    Me.frmTransactionLock.Enabled = True
    Me.frmTransactionLock.Form.trnEntryDate.ForeColor = 0
    Me.frmTransactionLock.Form.trnInternalRef.ForeColor = 0
    Me.frmTransactionLock.Form.trnPageCounter.ForeColor = 0
    Me.frmTransactionLock.Form.persID.ForeColor = 0
    Me.frmTransactionLock.Form.trnYear.ForeColor = 0
    Me.frmTransactionLock.Form.trnEntryDate.BorderColor = 0
    Me.frmTransactionLock.Form.trnInternalRef.BorderColor = 0
    Me.frmTransactionLock.Form.trnPageCounter.BorderColor = 0
    Me.frmTransactionLock.Form.persID.BorderColor = 0
    Me.frmTransactionLock.Form.trnYear.BorderColor = 0
    Me.cmdCommit.Enabled = True
End If
End Sub


Private Sub FnAllocAC(accID As Integer)
    
    If accID = 0 Then Exit Sub
    
    Dim rs As Recordset
    Dim strSQL As String

    strSQL = "SELECT nz(tblControl.ctrLYPLA,0) AS ctrLYPLA, nz(tblControl.ctrControlAcc,0) AS ctrControlAcc, nz(tblControl.ctrInvCR,0) AS ctrInvCR, nz(tblControl.ctrCRN_DR, 0) AS ctrCRN_DR, nz(tblControl.ctrDRN_CR, 0) AS ctrDRN_CR, nz(tblControl.ctrRcptDR,0) AS ctrRcptDR, nz(tblControl.ctrExpDR,0) AS ctrExpDR, nz(tblControl.ctrPaymtCR,0) AS ctrPaymtCR, nz(tblControl.ctrDeptsCR,0) AS ctrDeptsCR, nz(tblControl.ctrCshWithdrwDR,0) AS ctrCshWithdrwDR " & _
             "FROM tblControl;"
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    rs.MoveFirst
    If Forms.frmControl.ActiveControl.name <> "ctrLYPLA" _
    And (rs!ctrLYPLA = accID _
    Or Forms.frmControl.ctrLYPLA = accID) Then
        MsgBox FnIsErr(7011, Forms.frmControl.lblLYPLA.Caption), vbExclamation
        DoCmd.CancelEvent
        Exit Sub
    End If

'Checking if selected account was already selected somewhere else
    If Forms.frmControl.ActiveControl.name <> "ctrControlAcc" _
    And (rs!ctrControlAcc = accID _
    Or Forms.frmControl.ctrControlAcc = accID) Then
        MsgBox FnIsErr(7011, Forms.frmControl.lblControlAcc.Caption), vbExclamation
        DoCmd.CancelEvent
        Exit Sub
    End If

'    If Forms.frmControl.ActiveControl.Name <> "cboInvCR" _
'    And (rs!ctrInvCR = accID _
'    Or Forms.frmControl.cboInvCR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblInvCR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboCRN_DR" _
'    And (rs!ctrCRN_DR = accID _
'    Or Forms.frmControl.cboCRN_DR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblCRN_DR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboDRN_CR" _
'    And (rs!ctrDRN_CR = accID _
'    Or Forms.frmControl.cboDRN_CR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblDRN_CR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboRcptDR" _
'    And (rs!ctrRcptDR = accID _
'    Or Forms.frmControl.cboRcptDR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblRcptDR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboExpDR" _
'    And (rs!ctrExpDR = accID _
'    Or Forms.frmControl.cboExpDR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblExpDR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboPaymtCR" _
'    And (rs!ctrPaymtCR = accID _
'    Or Forms.frmControl.cboPaymtCR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblPaymtCR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboDeptsCR" _
'    And (rs!ctrDeptsCR = accID _
'    Or Forms.frmControl.cboDeptsCR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblDeptsCR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If
'
'    If Forms.frmControl.ActiveControl.Name <> "cboCshWithdrwDR" _
'    And (rs!ctrCshWithdrwDR = accID _
'    Or Forms.frmControl.cboCshWithdrwDR = accID) Then
'        MsgBox FnIsErr(7011, Forms.frmControl.lblCshWithdrwDR.Caption), vbExclamation
'        DoCmd.CancelEvent
'        Exit Sub
'    End If

End Sub
