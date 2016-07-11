Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =16440
    DatasheetFontHeight =10
    ItemSuffix =52
    Left =645
    Top =1830
    Right =16755
    Bottom =6705
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1093e6cbac7ae340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tblTransactionSub.* FROM tblTransactionSub; "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            Width =1701
        End
        Begin CommandButton
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =340
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =855
                    Top =60
                    Width =780
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label92"
                    Caption ="Account"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =3390
                    Top =60
                    Width =735
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label93"
                    Caption ="Date"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =5085
                    Top =60
                    Width =525
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label94"
                    Caption ="Ref."
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =13401
                    Top =60
                    Width =765
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label97"
                    Caption ="Debits"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =14595
                    Top =60
                    Width =765
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label98"
                    Caption ="Credits"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =8700
                    Top =60
                    Width =540
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label19"
                    Caption ="Note"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =4393
                    Top =60
                    Width =435
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label25"
                    Caption ="Vat"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Left =5685
                    Top =60
                    Width =1020
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label49"
                    Caption ="Doc Number."
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =6675
                    Top =60
                    Width =765
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label50"
                    Caption ="Doc Date"
                    FontName ="MS Sans Serif"
                End
            End
        End
        Begin Section
            Height =341
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =990
                    Height =256
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDate"
                    ControlSource ="trnsDate"
                    Format ="Short Date"
                    ValidationRule ="Month([trnsDate])=[Forms]![frmTransaction]![persID]"
                    ValidationText ="Wrong month in the transaction date (has to match period)."
                    DefaultValue ="=Forms(\"frmTransaction\").[trnEntryDate]"
                    FontName ="MS Sans Serif"
                    InputMask ="99/99/9999"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3119
                    Left =1185
                    Top =60
                    Width =2160
                    Height =256
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="accID2"
                    ControlSource ="accID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo, tblAccount.accSig"
                        "n, tblAccount.accVatable FROM tblAccount WHERE (((tblAccount.accStatus)=True)) O"
                        "RDER BY tblAccount.accName; "
                    ColumnWidths ="0;3402;852;0"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =1701
                    Left =10800
                    Top =60
                    Width =720
                    Height =256
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\"\"\"\"\";\"\";\"\";\"10\";\"6\""
                    Name ="crnID"
                    ControlSource ="crnID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCurrency.crnID, tblCurrency.crnShortName, tblCurrency.crnDefault, tblC"
                        "urrency.crnExchangeRate FROM tblCurrency ORDER BY tblCurrency.crnDefault, tblCur"
                        "rency.crnShortName; "
                    ColumnWidths ="0;567;0;1134"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =5010
                    Top =60
                    Width =630
                    Height =256
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="refID"
                    ControlSource ="refID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblReference.refID, tblReference.refShortName, tblReference.refName FROM "
                        "tblReference ORDER BY tblReference.refShortName; "
                    ColumnWidths ="0;567;1701"
                    OnEnter ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    LimitToList = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13589
                    Top =60
                    Width =465
                    Height =270
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsSign"
                    ControlSource ="trnsSign"
                    RowSourceType ="Value List"
                    RowSource ="C;D"
                    DefaultValue ="\"C\""
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12098
                    Top =66
                    Width =1080
                    Height =256
                    TabIndex =15
                    BackColor =10079487
                    ForeColor =-2147483640
                    Name ="trnsFAmount"
                    ControlSource ="trnsFAmount"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText ="Foren Amount can be only \"0\" or biger."
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x80000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13646
                    Top =60
                    Width =1080
                    Height =255
                    TabIndex =13
                    BackColor =10092543
                    ForeColor =-2147483640
                    Name ="trnsAmount"
                    ControlSource ="trnsAmount"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText =" Amount can be only \"0\" or biger."
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11531
                    Top =66
                    Width =555
                    Height =256
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsRate"
                    ControlSource ="trnsRate"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13703
                    Top =60
                    Width =210
                    Height =255
                    TabIndex =17
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsVAmount"
                    ControlSource ="trnsVAmount"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =855
                    Top =60
                    Width =255
                    Height =256
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnID"
                    ControlSource ="trnID"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =285
                    Top =60
                    Width =480
                    Height =256
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsID"
                    ControlSource ="trnsID"
                    FontName ="MS Sans Serif"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =68
                    Left =15555
                    Top =64
                    Width =186
                    Height =276
                    TabIndex =14
                    ForeColor =0
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000006000000060000000100180000000000780000000000000000000000 ,
                        0x0000000000000000000080000080000080000080000080000080000000008000 ,
                        0x0080000080000080000080000080000000008000008000008000008000008000 ,
                        0x0080000000008000008000008000008000008000008000000000800000800000 ,
                        0x8000008000008000008000000000800000800000800000800000800000800000
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
                    End
                    ControlTipText ="Next Record (Alt+N)"
                    Picture ="C:\\Documents and Settings\\d1\\Desktop\\d.bmp"
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7740
                    Top =60
                    Width =3045
                    Height =256
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsNote"
                    ControlSource ="trnsNote"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =13230
                    Top =64
                    Width =1134
                    Height =256
                    TabIndex =9
                    Name ="trnsDebits"
                    ControlSource ="trnsDebits"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText ="Value has to be larger or equal 0"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                End
                Begin TextBox
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =14415
                    Top =64
                    Width =1134
                    Height =256
                    TabIndex =18
                    Name ="trnsCredits"
                    ControlSource ="trnsCredits"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText ="Value has to be larger or equal 0"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =4365
                    Top =60
                    Width =630
                    Height =256
                    TabIndex =4
                    BoundColumn =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="trnsVAT"
                    ControlSource ="trnsVAT"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblVat.vatName, IIf(tblVAT.vatRate<0,0,tblVAT.vatRate) AS vatRate1, tblVa"
                        "t.vatRate, tblVat.vatID, tblVat.vatStartDate, tblVat.vatEndDate FROM tblVat WHER"
                        "E (((tblVat.vatStartDate)<=Date()) AND ((tblVat.vatEndDate)>=Date())) ORDER BY t"
                        "blVat.vatRate; "
                    ColumnWidths ="567;1701;0"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                    OnGotFocus ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =86
                    Left =15780
                    Top =64
                    Width =501
                    Height =277
                    TabIndex =19
                    Name ="trnsbtVAT"
                    Caption ="&VAT"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Add VAT entry"
                    UnicodeAccessKey =86
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13889
                    Top =56
                    Width =450
                    Height =255
                    TabIndex =20
                    ForeColor =-2147483640
                    Name ="trnsVID"
                    ControlSource ="trnsVID"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText =" Amount can be only \"0\" or biger."
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14796
                    Top =56
                    Width =450
                    Height =255
                    TabIndex =21
                    ForeColor =-2147483640
                    Name ="trnsIsVat"
                    ControlSource ="trnsIsVAT"
                    Format ="Standard"
                    ValidationRule =">=0"
                    ValidationText =" Amount can be only \"0\" or biger."
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =60
                    Width =1071
                    Height =256
                    TabIndex =6
                    Name ="docNo"
                    ControlSource ="docNo"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3119
                    Left =30
                    Top =60
                    Width =1140
                    Height =256
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"24\""
                    Name ="accID1"
                    ControlSource ="accID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName, tblAccount.accSig"
                        "n, tblAccount.accVatable FROM tblAccount WHERE (((tblAccount.accStatus)=True)) O"
                        "RDER BY tblAccount.accNo; "
                    ColumnWidths ="0;856;2268;0"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6735
                    Top =60
                    Width =990
                    Height =256
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDocDate"
                    ControlSource ="trnsDocDate"
                    Format ="Short Date"
                    FontName ="MS Sans Serif"
                    InputMask ="99/99/9999"
                End
            End
        End
        Begin FormFooter
            Height =453
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =5
                    Left =195
                    Top =15
                    Width =13263
                    Name ="Line100"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12638
                    Top =60
                    Width =1080
                    Height =225
                    BackColor =10079487
                    ForeColor =8388608
                    Name ="txtSumFAmount"
                    ControlSource ="=fsFAmount([trnID])"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8953
                    Top =60
                    Width =1080
                    Height =225
                    TabIndex =1
                    BackColor =10079487
                    ForeColor =8388608
                    Name ="txtSumAmount"
                    ControlSource ="=[txtSumCredits]-[txtSumDebits]"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8160
                    Top =60
                    Width =720
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label13"
                    Caption ="Balance:"
                    FontName ="MS Sans Serif"
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =75
                    Top =75
                    Width =5190
                    Height =302
                    TabIndex =2
                    Name ="cmdTrns"
                    Caption ="Add New Transaction"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10257
                    Top =60
                    Width =1080
                    Height =225
                    TabIndex =3
                    BackColor =10079487
                    ForeColor =8388608
                    Name ="txtSumDebits"
                    ControlSource ="=IIf(IsNumeric(fsCR_DB([trnID],\"trnsDebits\")),fsCR_DB([trnID],\"trnsDebits\"),"
                        "0)"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11391
                    Top =60
                    Width =1080
                    Height =225
                    TabIndex =4
                    BackColor =10079487
                    ForeColor =8388608
                    Name ="txtSumCredits"
                    ControlSource ="=IIf(IsNumeric(fsCR_DB([trnID],\"trnsCredits\")),fsCR_DB([trnID],\"trnsCredits\""
                        "),0)"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
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

Private Sub accID1_AfterUpdate()
If IsNoData(trnsSign) = True Then
    trnsSign.Value = accID1.Column(3)
End If
End Sub

Private Sub accID1_DblClick(Cancel As Integer)
DoCmd.OpenForm "frmAccount", , , , , , Me.Form.name
End Sub

Private Sub accID2_AfterUpdate()
If IsNoData(trnsSign) = True Then
    trnsSign.Value = accID2.Column(3)
End If
End Sub

Private Sub accID2_DblClick(Cancel As Integer)
DoCmd.OpenForm "frmAccount", , , , , , Me.Form.name
End Sub

Private Sub cmdDel_Click()
On Error GoTo RecordDeleted

If Forms("frmTransaction").Controls("cmdSave").Enabled = True Then
    If MsgBox("Do You want to Detele transaction?", vbOKCancel, "Delete Transaction") = vbOK Then
       If (IsNull(trnsID.Value) = False) And (trnsID.Value <> "") Then
          Dim lngtrnsID As Long
          lngtrnsID = trnsID
          accID1 = Empty
          trnsSign = "C"
          trnsVAT = Empty
          DoCmd.GoToRecord , , acGoTo
          DoCmd.SetWarnings False
          strSQL = "DELETE from tblTransactionSub WHERE trnsID=" & lngtrnsID
          FnLog (strSQL)
          DoCmd.RunSQL (strSQL)
          DoCmd.SetWarnings True
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
       End If
       Me.Requery
       Me.Refresh
    Else
    End If
    Me.Parent!HasChanged = True
End If

RecordDeleted:
      If Err.Number = 3167 Then
         Err.Clear
         Me.Requery
         Resume Next
     End If
End Sub

Private Sub cmdTrns_Click()
On Error GoTo Err_cmdAdd_Click
  Forms("frmTransaction").AllowAdditions = True
  Forms("frmTransaction").AllowEdits = True
  Me.AllowAdditions = True
  DoCmd.GoToRecord , , acNewRec
  accID1.SetFocus
  Forms("frmTransaction").Controls("cmdSave").Enabled = True
  Forms("frmTransaction").Controls("cmdExit").Enabled = False
  Forms("frmTransaction").Controls("cmdAdd").Enabled = False
  Forms("frmTransaction").Controls("cmdEdit").Enabled = False
Exit_cmdAdd_Click:
    Exit Sub
Err_cmdAdd_Click:
    MsgBox FnIsErr(Err.Number), vbExclamation
    Resume Exit_cmdAdd_Click
End Sub

Private Sub crnID_Change()
If IsNoData(crnID) = True Or crnID = 0 Then
   crnID = 0
Else

    trnsRate.Value = crnID.Column(3)
    If crnID.Column(2) = 0 Then
        trnsFAmount.Enabled = True
    Else
        trnsFAmount.Value = 0
        trnsFAmount.Enabled = False
    End If
End If

End Sub

Function tAmount(lngID As Long) As Double
Dim strSQL As String
strSQL = "SELECT Sum(IIf([trnsSign]='C',[trnsFAmount],-1*[trnsFAmount])) AS tAmount " _
         & "FROM tblTransactionSub " _
         & "WHERE (((tblTransactionSub.trnID)=1));"

Dim Db As Database
Dim rs As DAO.Recordset
Set Db = CurrentDb()
Set rs = Db.OpenRecordset(strSQL)
rs.MoveFirst
'rs (1)
'
End Function

Private Sub crnID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frmCurrency", , , , , , Me.Form.name
End Sub

Private Sub Form_AfterUpdate()
Me.Parent!HasChanged = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If IsNull(accID) Then
    MsgBox "The account has to be set.", vbExclamation
    DoCmd.CancelEvent
    Exit Sub
End If
End Sub

Private Sub Form_Current()
txtSumFAmount.Requery
txtSumAmount.Requery
refID.Requery
End Sub
Private Sub Form_Dirty(Cancel As Integer)
txtSumFAmount.Requery
txtSumAmount.Requery
End Sub
Private Sub Command22_Click()
On Error GoTo Err_Command22_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_Command22_Click:
    Exit Sub

Err_Command22_Click:
    MsgBox Err.Description
    Resume Exit_Command22_Click
    
End Sub

Private Sub refID_DblClick(Cancel As Integer)
DoCmd.OpenForm "frmReference", , , , , , "frmTransactionSub"
End Sub

Private Sub refID_Enter()
    Me.refID.Dropdown
End Sub

Private Sub trnsAmount_AfterUpdate()
    If (IsNoData(trnsFAmount) = True Or trnsFAmount.Value = 0) _
        And IsNoData(trnsAmount) = False _
        And IsNoData(trnsRate) = False Then
        trnsFAmount.Value = RoundDG(trnsAmount / trnsRate)  '--> modi
    End If
End Sub

Private Sub trnsbtVAT_Click()
    If trnsIsVat = True Then
        MsgBox "This is a VAT entry. You cannot add a VAT entry for the existing VAT entry.", vbOKOnly, "VAT"
        Exit Sub
    End If
    
    If trnsVAT < 0 Then
        MsgBox "This entry doesn't generate VAT. Check selected VAT.", vbOKOnly, "VAT"
        Exit Sub
    End If
    
    If accID1.Column(4) = False Then
        MsgBox "Selected account is not a vatable account.", vbOKOnly, "VAT"
        Exit Sub
    End If
    
    If trnsVAT.ListIndex <= 0 Then
        trnsVAT.SetFocus
        Exit Sub
    End If
    
'Checks if there is a VAT entry for that transaction
    Dim strSQL As String
    strSQL = "SELECT trnsID " _
         & "FROM tblTransactionSub " _
         & "WHERE (((tblTransactionSub.trnsVID)=" & trnsID & ")) AND " _
         & "      tblTransactionSub.trnsIsVat = TRUE;"

    Dim Db As Database
    Dim rs As DAO.Recordset
    Set Db = CurrentDb()
    Set rs = Db.OpenRecordset(strSQL)
    If rs.RecordCount > 0 Then
        MsgBox "This transaction has already VAT entry.", vbOKOnly, "VAT"
        Exit Sub
    End If
'Inserts VAT entry
    If IsNull(trnsVAT.Value) = False Then
        If Forms("frmTransaction").Controls("cmdSave").Enabled = True Then
            If (IsNull(trnsID.Value) = False) And (trnsID.Value <> "") Then
                DoCmd.SetWarnings False
                strSQL = "INSERT INTO " _
                       & "tblTransactionSub(trnID, accID, trnsDate, trnsVAT, trnsVID, trnsIsVat, crnID, refID, desID, trnsSign, trnsFAmount, trnsAmount, trnsRate, trnsVAmount, trnsNote, trnsDebits, trnsCredits) " _
                       & "SELECT trnID, " _
                       & "       IIf(trnsSign='D', vatinput, vatoutput)as accid, " _
                       & "       trnsDate, " _
                       & "       0 as trnsVAT, " _
                       & "       trnsID as trnsVID,  " _
                       & "       1 as trnsIsVat, " _
                       & "       crnID, " _
                       & "       refID, " _
                       & "       desID, " _
                       & "       trnsSign, " _
                       & "       trnsFAmount, " _
                       & "       trnsAmount, " _
                       & "       trnsRate, " _
                       & "       trnsVAmount, trnsNote, IIf(trnsSign='D',trnsVAmount,0) as trnsDebits, IIf(trnsSign='C',trnsVAmount,0) as trnsCredits FROM tblTransactionSub INNER JOIN tblVat ON tblTransactionSub.trnsVID = tblVat.vatID WHERE trnsID = " & trnsID.Value
                FnLog (strSQL)
                DoCmd.RunSQL (strSQL)
                DoCmd.SetWarnings True
            End If
            Me.Requery
            Me.Refresh
            DoCmd.GoToRecord , , acLast
        Else
        End If
    End If
End Sub
Private Sub trnsCredits_AfterUpdate()
'Flag used for automatic update of opening balances for a next year if any changes have been
'made to current year
'Me.Parent!HasChanged = True
End Sub

Private Sub trnsDebits_AfterUpdate()
'    If (trnsCredits > 0) And (trnsDebits > 0) Then
'        MsgBox "Value allowed only in Debits or Credits", vbOKOnly, "One field value"
'        trnsDebits.SetFocus
'        Exit Sub
'    End If
'    If (trnsCredits = 0) And (trnsDebits > 0) Then
'        trnsSign.Value = "D"
'        trnsAmount.Value = trnsDebits.Value
'    End If

'Flag used for automatic update of opening balances for a next year if any changes have been
'made to current year
'Me.Parent!HasChanged = True
End Sub

Private Sub trnsDebits_DblClick(Cancel As Integer)
    If IsNull(trnsID) Then
        Exit Sub
    End If
    If Me.AllowEdits = True Then
        DoCmd.OpenForm "frmTransactionSingle", , , "trnsID=" & trnsID
    Else
        DoCmd.OpenForm "frmTransactionSingle", , , "trnsID=" & trnsID, acFormReadOnly
    End If
End Sub


Private Sub trnsCredits_DblClick(Cancel As Integer)
    If Me.AllowEdits = True Then
        DoCmd.OpenForm "frmTransactionSingle", , , "trnsID=" & trnsID
    Else
        DoCmd.OpenForm "frmTransactionSingle", , , "trnsID=" & trnsID, acFormReadOnly
    End If
End Sub

Private Sub trnsDebits_Exit(Cancel As Integer)
    If (trnsCredits > 0) And (trnsDebits > 0) Then
        MsgBox "Value allowed only in Debits or Credits", vbOKOnly, "One field value"
        DoCmd.CancelEvent
        Exit Sub
    End If
    If (trnsCredits = 0) And (trnsDebits > 0) Then
        trnsSign.Value = "D"
        trnsAmount = trnsDebits
        trnsbtVAT.Enabled = True
        If IsNull(trnsVAT) Then
            trnsVAT.SetFocus
            Exit Sub
        Else
            If (IsNull(trnsVAT) = False And trnsVAT <> 0) Or trnsVAT <> Empty Then
                trnsVID = trnsVAT.Column(3)
            End If
            trnsVAmount = trnsDebits * (trnsVAT / 100)
        End If
        If IsNull(accID1) Then
            accID1.SetFocus
        Else
            Me.Refresh
        End If
    End If
End Sub
Private Sub trnsCredits_Exit(Cancel As Integer)
    If (trnsDebits.Value > 0) And (trnsCredits.Value > 0) Then
        MsgBox "Value allowed only in Debits or Credits", vbOKOnly, "One field value"
        DoCmd.CancelEvent
        Exit Sub
    End If
    If (trnsDebits = 0) And (trnsCredits > 0) Then
        trnsSign.Value = "C"
        trnsAmount.Value = trnsCredits.Value
        trnsbtVAT.Enabled = True
        If IsNull(trnsVAT) Then
            trnsVAT.SetFocus
            Exit Sub
        Else
            If (IsNull(trnsVAT) = False And trnsVAT <> 0) Or trnsVAT <> Empty Then
                trnsVID = trnsVAT.Column(3)
            End If
            trnsVAmount = trnsCredits * (trnsVAT / 100)
        End If
        If IsNull(accID1) Then
            accID1.SetFocus
        Else
            Me.Refresh
        End If
    End If
End Sub


Private Sub trnsVAT_DblClick(Cancel As Integer)
DoCmd.OpenForm "frmVat", , , , , , Me.Form.name
End Sub
Private Sub Command40_Click()
On Error GoTo Err_Command40_Click

    Dim stDocName As String

    stDocName = "tbHelp"
    DoCmd.RunMacro stDocName

Exit_Command40_Click:
    Exit Sub

Err_Command40_Click:
    MsgBox Err.Description
    Resume Exit_Command40_Click
    
End Sub

Private Sub trnsVAT_Exit(Cancel As Integer)
On Error GoTo Err_trnsVAT_Exit

    If trnsIsVat = False Then
        If (IsNull(trnsVAT) = False And trnsVAT <> 0) Or trnsVAT <> Empty Then
            trnsVID.Value = CInt(trnsVAT.Column(3))
            If trnsDebits.Value <> 0 Then
                trnsVAmount.Value = trnsDebits * (trnsVAT / 100)
            ElseIf trnsCredits.Value <> 0 Then
                trnsVAmount.Value = trnsCredits * (trnsVAT / 100)
            End If
            trnsVAT.Requery
            'Me.Refresh
        End If
    End If

Exit_trnsVAT_Exit:
    Exit Sub

Err_trnsVAT_Exit:
    If Err.Number <> 2001 Then
        MsgBox FnIsErr(Err.Number), vbExclamation
    End If
    Resume Exit_trnsVAT_Exit
End Sub


Private Sub trnsVAT_GotFocus()
    If trnsIsVat = True Then
        MsgBox "This is a VAT entry. You cannot set VAT rate for the VAT entry.", vbOKOnly, "VAT"
        refID.SetFocus
        DoCmd.CancelEvent
        Exit Sub
    End If
End Sub
