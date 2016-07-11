Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =27
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9014
    DatasheetFontHeight =10
    ItemSuffix =19
    Left =3090
    Top =1020
    Right =12105
    Bottom =9060
    DatasheetGridlinesColor =12632256
    Filter ="trnsID=175"
    RecSrcDt = Begin
        0xaddcbc97d377e340
    End
    RecordSource ="SELECT tblTransaction.trnEntryDate, tblTransaction.trnInternalRef, tblTransactio"
        "n.trnPageCounter, tblTransaction.trnYear, tblTransaction.trnLock, tblTransaction"
        "Sub.* FROM tblTransactionSub INNER JOIN tblTransaction ON tblTransactionSub.trnI"
        "D=tblTransaction.trnID; "
    Caption ="Transaction Details"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
    AfterFinalRender ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1417
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =226
                    Top =284
                    Width =8629
                    Height =1086
                    Name ="Frame125"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =368
                            Top =113
                            Width =975
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label126"
                            Caption ="Page Details"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    Left =1366
                    Top =396
                    Width =1485
                    TabIndex =1
                    BackColor =-2147483633
                    ForeColor =128
                    Name ="trnID"
                    ControlSource ="trnID"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =301
                            Top =396
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label115"
                            Caption ="Page ID"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    Left =1366
                    Top =679
                    Width =1485
                    TabIndex =2
                    BackColor =-2147483633
                    Name ="trnEntryDate"
                    ControlSource ="trnEntryDate"
                    Format ="Short Date"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =301
                            Top =679
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label95"
                            Caption ="EntryDate"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    Left =1366
                    Top =963
                    Width =1485
                    TabIndex =3
                    BackColor =-2147483633
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =226
                            Top =963
                            Width =1080
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label96"
                            Caption ="InternalRef."
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    Left =4657
                    Top =680
                    Width =1440
                    TabIndex =4
                    BackColor =-2147483633
                    Name ="trnYear"
                    ControlSource ="trnYear"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =3967
                            Top =680
                            Width =630
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label98"
                            Caption ="Year:"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    Left =4666
                    Top =987
                    Width =1440
                    TabIndex =5
                    BackColor =-2147483633
                    Name ="trnPageCounter"
                    ControlSource ="trnPageCounter"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =3376
                            Top =987
                            Width =1230
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label97"
                            Caption ="PageCounter:"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6800
                    Top =340
                    Width =1290
                    Height =255
                    FontWeight =700
                    TabIndex =6
                    BackColor =-2147483633
                    ForeColor =128
                    Name ="txtReadOnly"
                    ControlSource ="=IIf([AllowEdits]=False,\"Read Only\",\"\")"
                    Format ="Standard"
                    FontName ="MS Sans Serif"

                End
            End
        End
        Begin Section
            Height =4535
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1449
                    Top =975
                    Width =1380
                    Height =255
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDate"
                    ControlSource ="trnsDate"
                    Format ="Short Date"
                    DefaultValue ="=Forms(\"frmTransaction\").[trnEntryDate]"
                    FontName ="MS Sans Serif"
                    InputMask ="00/00/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =441
                            Top =975
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label124"
                            Caption ="Date"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3119
                    Left =1449
                    Top =645
                    Width =3645
                    Height =270
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="accID"
                    ControlSource ="accID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo, tblAccount.accSig"
                        "n FROM tblAccount WHERE (((tblAccount.accStatus)=True)) ORDER BY tblAccount.accN"
                        "ame; "
                    ColumnWidths ="0;2268;855;0"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =441
                            Top =645
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label113"
                            Caption ="Account"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =1701
                    Left =1449
                    Top =3086
                    Width =780
                    Height =270
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="crnID"
                    ControlSource ="crnID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCurrency.crnID, tblCurrency.crnShortName, tblCurrency.crnDefault, tblC"
                        "urrency.crnExchangeRate FROM tblCurrency ORDER BY tblCurrency.crnDefault, tblCur"
                        "rency.crnShortName;"
                    ColumnWidths ="0;567;0;1134"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =441
                            Top =3116
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label116"
                            Caption ="Currency"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =1449
                    Top =1275
                    Width =780
                    Height =270
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="refID"
                    ControlSource ="refID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblReference.refID, tblReference.refShortName, tblReference.refName FROM "
                        "tblReference ORDER BY tblReference.refShortName; "
                    ColumnWidths ="0;567;1701"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =441
                            Top =1275
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label117"
                            Caption ="Reference"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =2268
                    Left =1449
                    Top =2745
                    Width =780
                    Height =270
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsSign"
                    ControlSource ="trnsSign"
                    RowSourceType ="Value List"
                    RowSource ="C;D"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =441
                            Top =2745
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label118"
                            Caption ="Sign"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3717
                    Top =3101
                    Width =1470
                    Height =255
                    TabIndex =12
                    BackColor =10079487
                    ForeColor =-2147483640
                    Name ="trnsFAmount"
                    ControlSource ="trnsFAmount"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText ="Foren Amount can be only \"0\" or biger."
                    OnExit ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1449
                    Top =3450
                    Width =1470
                    Height =255
                    TabIndex =15
                    BackColor =10092543
                    ForeColor =-2147483640
                    Name ="trnsAmount"
                    ControlSource ="trnsAmount"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText =" Amount can be only \"0\" or biger."
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =441
                            Top =3450
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label120"
                            Caption ="Amount"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2936
                    Top =3101
                    Width =750
                    Height =255
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsRate"
                    ControlSource ="trnsRate"
                    Format ="Standard"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2306
                            Top =3116
                            Width =570
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label119"
                            Caption ="Rate"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1449
                    Top =3945
                    Width =1470
                    Height =255
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsVAmount"
                    ControlSource ="trnsVAmount"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =435
                            Top =3945
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label121"
                            Caption ="Vatable A."
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    Left =1447
                    Top =359
                    Width =1965
                    Height =255
                    TabIndex =1
                    BackColor =-2147483633
                    ForeColor =128
                    Name ="trnsID"
                    ControlSource ="trnsID"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =727
                            Top =359
                            Width =660
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label114"
                            Caption ="ID"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1449
                    Top =2295
                    Width =5910
                    Height =256
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsNote"
                    ControlSource ="trnsNote"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =441
                            Top =2295
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label123"
                            Caption ="Note"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin OptionGroup
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =223
                    Top =171
                    Width =8689
                    Height =4311
                    Name ="Frame127"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =425
                            Width =1440
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label128"
                            Caption ="Transaction Details"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =4409
                    Top =1303
                    Width =780
                    Height =270
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="desID"
                    ControlSource ="desID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblDescription.desID, tblDescription.desShortName, tblDescription.desName"
                        " FROM tblDescription ORDER BY tblDescription.desShortName; "
                    ColumnWidths ="0;567;1701"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =3344
                            Top =1303
                            Width =1005
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label131"
                            Caption ="Description"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1449
                    Top =1587
                    Width =750
                    Height =255
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsVAT"
                    ControlSource ="trnsVAT"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText ="VATt can be only \"0\" or biger."
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =876
                            Top =1587
                            Width =570
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label7"
                            Caption ="VAT"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3049
                    Top =3431
                    Width =907
                    Height =271
                    TabIndex =17
                    BackColor =-2147483633
                    Name ="txtAmtDesc"
                    ControlSource ="=IIf([trnsSign]=\"D\",\"Debits\",\"Credits\")"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5265
                    Top =3116
                    TabIndex =13
                    Name ="trnsDebits"
                    ControlSource ="trnsDebits"
                    Format ="Standard"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7022
                    Top =3116
                    TabIndex =14
                    Name ="trnsCredits"
                    ControlSource ="trnsCredits"
                    Format ="Standard"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1449
                    Top =1995
                    Width =1380
                    Height =255
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDocDate"
                    ControlSource ="trnsDocDate"
                    Format ="Short Date"
                    DefaultValue ="=Forms(\"frmTransaction\").[trnEntryDate]"
                    FontName ="MS Sans Serif"
                    InputMask ="00/00/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =246
                            Top =1995
                            Width =1200
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label15"
                            Caption ="Document date"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =5910
                    Top =2816
                    Width =1035
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label16"
                    Caption ="Debit Amount"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =7698
                    Top =2816
                    Width =1005
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label17"
                    Caption ="Credit Amount"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =4170
                    Top =2816
                    Width =1005
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label18"
                    Caption ="F. Amount"
                    FontName ="MS Sans Serif"
                End
            End
        End
        Begin FormFooter
            Height =566
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    Left =225
                    Top =53
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
                        0x000301000000000000000000
                    End
                    ControlTipText ="First Record (Alt+M)"
                    UnicodeAccessKey =109

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =47
                    Left =1485
                    Top =53
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
                        0x000301000000000000000000
                    End
                    ControlTipText ="Last Record (Alt+/)"
                    UnicodeAccessKey =47

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =46
                    Left =1065
                    Top =53
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
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record (Alt+.)"
                    UnicodeAccessKey =46

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =44
                    Left =645
                    Top =53
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
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record (Alt+,)"
                    UnicodeAccessKey =44

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =7166
                    Top =45
                    Width =1080
                    Height =360
                    TabIndex =4
                    ForeColor =0
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Close Current Form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub cmdExit_Click()
    DoCmd.Close
End Sub

Private Sub crnID_Change()
    trnsRate.Value = crnID.Column(3)
    If crnID.Column(2) = 0 Then
        trnsFAmount.Enabled = True
    Else
        trnsFAmount.Value = 0
        trnsFAmount.Enabled = False
    End If
End Sub

Private Sub crnID_DblClick(Cancel As Integer)
'    Me.Modal = False
    DoCmd.OpenForm "frmCurrency", , , , , , Me.Form.name
End Sub

Private Sub desID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frmDescription", , , , , , Me.Form.name
End Sub

Private Sub Form_AfterFinalRender(ByVal drawObject As Object)
trnsRate.Value = crnID.Column(3)
End Sub

Private Sub Form_Current()
If FnIsLoad("frmTransaction") Then trnsSign.SetFocus

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
'trnsRate.Value = CSng(crnID.Column(3))
   

End Sub

Private Sub refID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frmReference", , , , , , Me.Form.name
End Sub

Private Sub trnsAmount_AfterUpdate()
    If (IsNoData(trnsFAmount) = True Or trnsFAmount.Value = 0) _
        And IsNoData(trnsAmount) = False _
        And IsNoData(trnsRate) = False Then
        trnsFAmount.Value = RoundDG(trnsAmount / trnsRate) '-->modi
    End If
End Sub

Private Sub trnsFAmount_Change()
If IsNoData(trnsRate) Then trnsRate = 0
If IsNoData(trnsFAmount) Then trnsFAmount = 0
trnsAmount = RoundDG(trnsRate * trnsFAmount)
If trnsSign = "C" Then
   trnsDebits = 0
   trnsCredits = trnsAmount
Else
   trnsCredits = 0
   trnsDebits = trnsAmount
End If
End Sub

Private Sub trnsFAmount_Exit(Cancel As Integer)
If IsNoData(trnsRate) Then trnsRate = 0
If IsNoData(trnsFAmount) Then trnsFAmount = 0
trnsAmount = RoundDG(trnsRate * trnsFAmount)
If trnsSign = "C" Then
   trnsDebits = 0
   trnsCredits = trnsAmount
Else
   trnsCredits = 0
   trnsDebits = trnsAmount
End If
End Sub

Private Sub trnsSign_Change()
If trnsSign = "C" Then
   trnsDebits = 0
   trnsCredits = trnsAmount
Else
   trnsCredits = 0
   trnsDebits = trnsAmount
End If

End Sub
