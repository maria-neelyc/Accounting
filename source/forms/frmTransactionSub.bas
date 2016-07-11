Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15023
    DatasheetFontHeight =10
    ItemSuffix =55
    Left =165
    Top =1590
    Right =15615
    Bottom =5670
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x839ef4a82f9ae340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tblTransactionSub.*, tblTransactionSub.trnsID FROM tblTransactionSub ORDE"
        "R BY tblTransactionSub.trnsID; "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
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
                    Left =3570
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
                    Left =5265
                    Top =60
                    Width =570
                    Height =210
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
                    Left =11436
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
                    Left =12630
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
                    Left =8880
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
                    Left =4573
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
                    OverlapFlags =85
                    Left =5895
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
                    OverlapFlags =85
                    Left =7140
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
            Height =396
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3465
                    Top =60
                    Width =990
                    Height =255
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDate"
                    ControlSource ="trnsDate"
                    Format ="Short Date"
                    ValidationRule ="Month([trnsDate])=[Forms]![frmTransaction]![persID]"
                    ValidationText ="Wrong month in the transaction date (has to match period)."
                    DefaultValue ="=CopyLastTrnsField(\"trnsDate\",[trnID],Forms(\"frmTransaction\").[trnEntryDate]"
                        ")"
                    FontName ="Arial Greek"
                    InputMask ="99/99/9999"
                    OnKeyPress ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3119
                    Left =1260
                    Top =60
                    Width =2160
                    Height =270
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="accID2"
                    ControlSource ="accID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo, tblAccount.accSig"
                        "n, tblAccount.accVatable FROM tblAccount WHERE (((tblAccount.accStatus)=True)) O"
                        "RDER BY tblAccount.accName; "
                    ColumnWidths ="0;3402;852;0"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial Greek"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =1701
                    Left =13589
                    Top =60
                    Width =705
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
                    DefaultValue ="=1"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"
                    LayoutCachedLeft =13589
                    LayoutCachedTop =60
                    LayoutCachedWidth =14294
                    LayoutCachedHeight =330
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =5235
                    Top =60
                    Width =630
                    Height =285
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
                    DefaultValue ="=CopyLastTrnsField(\"refID\",[trnID])"
                    FontName ="Arial Greek"
                    LimitToList = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =255
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
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13759
                    Top =60
                    Width =1080
                    Height =255
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

                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000080000000ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
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
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13816
                    Top =60
                    Width =555
                    Height =255
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
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsID"
                    ControlSource ="tblTransactionSub.trnsID"
                    FontName ="MS Sans Serif"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =68
                    Left =13590
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
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record (Alt+D)"
                    Picture ="d.bmp"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8100
                    Top =64
                    Width =3045
                    Height =255
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsNote"
                    ControlSource ="trnsNote"
                    FontName ="Arial Greek"
                    OnKeyPress ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =11265
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
                    FontName ="Arial Greek"

                End
                Begin TextBox
                    OverlapFlags =119
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =12450
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
                    FontName ="Arial Greek"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =4500
                    Top =60
                    Width =630
                    Height =285
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
                        "blVat.vatRate;"
                    ColumnWidths ="567;1701;0"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Arial Greek"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =86
                    Left =13815
                    Top =64
                    Width =501
                    Height =277
                    TabIndex =19
                    Name ="trnsbtVAT"
                    Caption ="&VAT"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add VAT entry (ALT + V)"
                    UnicodeAccessKey =86

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14385
                    Top =64
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
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14550
                    Top =64
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
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =5910
                    Top =60
                    Width =1071
                    Height =256
                    TabIndex =6
                    Name ="docNo"
                    ControlSource ="docNo"
                    DefaultValue ="=CopyLastTrnsField(\"docNo\",[trnID])"
                    FontName ="Arial Greek"
                    OnKeyPress ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3119
                    Left =60
                    Top =60
                    Width =1140
                    Height =270
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
                    FontName ="Arial Greek"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7035
                    Top =60
                    Width =990
                    Height =255
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDocDate"
                    ControlSource ="trnsDocDate"
                    Format ="Short Date"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="=CopyLastTrnsField(\"trnsDocDate\",[trnID])"
                    FontName ="Arial Greek"
                    InputMask ="99/99/9999"
                    OnKeyPress ="[Event Procedure]"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextFontCharSet =161
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14340
                    Top =60
                    Width =600
                    Height =255
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsPostDate"
                    ControlSource ="trnsPostDate"
                    Format ="Short Date"
                    DefaultValue ="=Now()"
                    FontName ="Arial Greek"
                    InputMask ="99/99/9999"

                    LayoutCachedLeft =14340
                    LayoutCachedTop =60
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =315
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

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
          'DoCmd.GoToRecord , , acGoTo
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
    trnsRate.Value = crnID.Column(3)
    If crnID.Column(2) = 0 Then
        trnsFAmount.Enabled = True
    Else
        trnsFAmount.Value = 0
        trnsFAmount.Enabled = False
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

Private Sub docNo_KeyPress(KeyAscii As Integer)
    'Copy the above description
    If Len(docNo.Text) = 0 And KeyAscii = Asc("?") Then
        'strSQL = DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID)
        docNo = DLookup("docNo", "tblTransactionSub", "trnsID = " & Nz(DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID), 0))
        DoCmd.CancelEvent
    End If
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38 'Up Arrow
            If Not Me.Recordset.BOF Then
                Me.Recordset.MovePrevious
            End If
        Case 40 'Down Arrow
            If Not Me.Recordset.EOF Then
                Me.Recordset.MoveNext
            End If
    End Select
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
        trnsFAmount.Value = RoundDG(trnsAmount / trnsRate) '--?modi
    End If
End Sub

Private Sub trnsbtVAT_Click()
    DoCmd.RunCommand acCmdSaveRecord
    If trnsIsVat = True Then
        MsgBox "This is a VAT entry. You cannot add a VAT entry for the existing VAT entry.", vbOKOnly, "VAT"
        Exit Sub
    End If
    
    If trnsVAT < 0 _
    Or DLookup("vatInput", "tblVAT", "vatID = " & trnsVID) = 0 Then
        MsgBox "This entry doesn't generate VAT. Check selected VAT.", vbOKOnly, "VAT"
        Exit Sub
    End If
    
    If accID1.Column(4) = False Then
        MsgBox "Selected account is not a vatable account.", vbOKOnly, "VAT"
        Exit Sub
    End If
    
    If trnsVAT.ListIndex < 0 Then
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
                
                strSQL = "INSERT INTO tblTransactionSub(trnID, accID, trnsDate, trnsVAT, trnsVID, trnsIsVat, crnID, refID, desID, trnsSign, trnsFAmount, trnsAmount, trnsRate, trnsVAmount, trnsNote, trnsDebits, trnsCredits, trnsDocDate) " & _
                         "SELECT trnID, IIf(trnsSign='D', vatinput, vatoutput)as accid, trnsDate, 0 as trnsVAT, trnsID as trnsVID, 1 as trnsIsVat, crnID, refID, desID, trnsSign, trnsFAmount, trnsAmount, trnsRate, trnsVAmount, trnsNote, IIf(trnsSign='D',trnsVAmount,0) as trnsDebits, IIf(trnsSign='C',trnsVAmount,0) as trnsCredits, trnsDocDate " & _
                         "FROM tblTransactionSub INNER JOIN tblVat ON tblTransactionSub.trnsVID = tblVat.vatID WHERE trnsID = " & trnsID.Value
                FnLog (strSQL)
                DoCmd.RunSQL (strSQL)
'                DoCmd.RunSQL "INSERT INTO " _
'                            & "tblTransactionSub(trnID, accID, trnsDate, trnsVAT, trnsVID, trnsIsVat, crnID, refID, desID, trnsSign, trnsFAmount, trnsAmount, trnsRate, trnsVAmount, trnsNote, trnsDebits, trnsCredits, trnsDocDate) " _
'                            & "SELECT trnID, " _
'                            & "       IIf(trnsSign='D', vatinput, vatoutput)as accid, " _
'                            & "       trnsDate, " _
'                            & "       0 as trnsVAT, " _
'                            & "       trnsID as trnsVID,  " _
'                            & "       1 as trnsIsVat, " _
'                            & "       crnID, " _
'                            & "       refID, " _
'                            & "       desID, " _
'                            & "       trnsSign, " _
'                            & "       trnsFAmount, " _
'                            & "       trnsAmount, " _
'                            & "       trnsRate, " _
'                            & "       trnsVAmount, " _
'                            & "       trnsNote, " _
'                            & "       IIf(trnsSign='D',trnsVAmount,0) as trnsDebits, " _
'                            & "       IIf(trnsSign='C',trnsVAmount,0) as trnsCredits, " _
'                            & "       trnsDocDate " _
'                            & "       FROM tblTransactionSub INNER JOIN tblVat ON tblTransactionSub.trnsVID = tblVat.vatID WHERE trnsID = " & trnsID.Value
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

Private Sub trnsDate_KeyPress(KeyAscii As Integer)
    'Copy the above description
    If IsNoData(trnsDate) And KeyAscii = Asc("?") Then
        'strSQL = DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID)
        trnsDate = DLookup("trnsDate", "tblTransactionSub", "trnsID = " & Nz(DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID), 0))
        DoCmd.CancelEvent
    End If
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
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.OpenForm "frmTransactionSingle", , , "trnsID=" & trnsID
        
    Else
        DoCmd.OpenForm "frmTransactionSingle", , , "trnsID=" & trnsID, acFormReadOnly
    End If
End Sub


Private Sub trnsCredits_DblClick(Cancel As Integer)
    If IsNull(trnsID) Then
        Exit Sub
    End If

    If Me.AllowEdits = True Then
        DoCmd.RunCommand acCmdSaveRecord
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
            trnsVAmount = RoundDG(trnsDebits * (trnsVAT / 100))
            trnsRate.Value = crnID.Column(3)                   '19/11/2009
                If (IsNoData(trnsFAmount) = True Or trnsFAmount.Value = 0) _
                    And IsNoData(trnsAmount) = False _
                    And IsNoData(trnsRate) = False Then
                    trnsFAmount.Value = RoundDG(trnsAmount / trnsRate) '--?modi
                End If
                'trnsFAmount.Value = RoundDG(trnsAmount / trnsRate) '19/11/2009 ->modi
        End If
        
        'update VAT entry if exists
        Dim rs As Recordset
    
        strSQL = "SELECT trnsID FROM tblTransactionSub WHERE trnsIsVat = True AND trnsVID = " & trnsID
        Set rs = CurrentDb.OpenRecordset(strSQL)
        If rs.RecordCount > 0 Then
            strSQL = "UPDATE tblTransactionSub " & _
                     "SET trnsDebits = " & Round(trnsDebits * (trnsVAT / 100), 2) & ", " & _
                     "trnsVAmount = " & Round(trnsDebits * (trnsVAT / 100), 2) & " " & _
                    "WHERE trnsID = " & rs!trnsID
            DoCmd.SetWarnings False
            FnLog (strSQL)
            DoCmd.RunSQL (strSQL)
            DoCmd.SetWarnings True
        End If
        
        rs.Close
        Set rs = Nothing
        
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
            trnsVAmount = RoundDG(trnsCredits * (trnsVAT / 100))
            trnsRate.Value = crnID.Column(3)                   '19/11/2009
                If (IsNoData(trnsFAmount) = True Or trnsFAmount.Value = 0) _
                    And IsNoData(trnsAmount) = False _
                    And IsNoData(trnsRate) = False Then
                    trnsFAmount.Value = RoundDG(trnsAmount / trnsRate) '--?modi
                End If
                'trnsFAmount.Value = RoundDG(trnsAmount / trnsRate) '19/11/2009 -->modi
        End If
        
        'update VAT entry if exists
        Dim rs As Recordset

        strSQL = "SELECT trnsID FROM tblTransactionSub WHERE trnsIsVat = True AND trnsVID = " & trnsID
        Set rs = CurrentDb.OpenRecordset(strSQL)
        If rs.RecordCount > 0 Then
            strSQL = "UPDATE tblTransactionSub " & _
                     "SET trnsCredits = " & Round(trnsCredits * (trnsVAT / 100), 2) & ", " & _
                     "trnsVAmount = " & Round(trnsCredits * (trnsVAT / 100), 2) & " " & _
                    "WHERE trnsID = " & rs!trnsID
            DoCmd.SetWarnings False
            FnLog (strSQL)
            DoCmd.RunSQL (strSQL)
            DoCmd.SetWarnings True
        End If

        rs.Close
        Set rs = Nothing
        
        If IsNull(accID1) Then
            accID1.SetFocus
        Else
            Me.Refresh
        End If
    End If
End Sub


Private Sub trnsDocDate_Exit(Cancel As Integer)
On Error GoTo Err_trnsDocDate_Exit

'Validate if the doc no is from the period
'If IsNull(trnsDocDate) = False Then
'    If Month(trnsDocDate) <> Forms("frmTransaction").Controls("persID").Value Then
'        MsgBox FnIsErr(7010), vbExclamation
'        DoCmd.CancelEvent
'    End If
'End If

Exit_trnsDocDate_Exit:
    Exit Sub

Err_trnsDocDate_Exit:
    If Err.Number <> 2001 Then
        MsgBox FnIsErr(Err.Number), vbExclamation
    End If
    Resume Exit_trnsDocDate_Exit
End Sub

Private Sub trnsDocDate_KeyPress(KeyAscii As Integer)

    'Copy the above description
    If IsNoData(trnsDocDate) And KeyAscii = Asc("?") Then
        'strSQL = DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID)
        trnsDocDate = DLookup("trnsDocDate", "tblTransactionSub", "trnsID = " & Nz(DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID), 0))
        DoCmd.CancelEvent
    End If
End Sub

Private Sub trnsNote_KeyPress(KeyAscii As Integer)
        
    'Copy the account name
    If Len(trnsNote.Text) = 0 And KeyAscii = Asc(" ") Then
        trnsNote = Me.accID2.Column(1)
        DoCmd.CancelEvent
    End If
    'Copy the above description
    If Len(trnsNote.Text) = 0 And KeyAscii = Asc("?") Then
        'strSQL = DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID)
        trnsNote = DLookup("trnsNote", "tblTransactionSub", "trnsID = " & Nz(DLookup("Max(trnsID)", "tblTransactionSub", "trnID = " & trnID & " AND trnsID < " & trnsID), 0))
        DoCmd.CancelEvent
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
                trnsVAmount.Value = RoundDG(trnsDebits * (trnsVAT / 100))
            ElseIf trnsCredits.Value <> 0 Then
                trnsVAmount.Value = RoundDG(trnsCredits * (trnsVAT / 100))
            End If
            trnsVAT.Requery
            'Me.Refresh
        End If
    End If
    If IsNoData(trnsVAT) Then
      trnsVAT.Value = 0
      trnsVID.Value = 0
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

'Private Sub trnsVAT_NotInList(NewData As String, Response As Integer)
'If IsNoData(trnsVAT.Value) Then
'    trnsVAT.Value = 0
'    trnsVID.Value = 0
'    refID.SetFocus
'    DoCmd.CancelEvent
'End If
'End Sub
