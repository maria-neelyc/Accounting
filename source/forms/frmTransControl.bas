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
    Width =9975
    DatasheetFontHeight =10
    ItemSuffix =365
    Left =6480
    Top =1752
    Right =17532
    Bottom =9048
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xacc8fcb3f07de440
    End
    RecordSource ="SELECT tblControl.*, sysMyCo.smycoName, sysMyCo.smycoPath, sysMyCo.smycoYear FRO"
        "M tblControl INNER JOIN sysMyCo ON tblControl.ctrDBID=sysMyCo.smycoID;"
    Caption ="Transaction Control"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa4050000a2050000a2050000a205000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionGroup
            SpecialEffect =3
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Subform
            SpecialEffect =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Tab
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Page
            Width =1701
            Height =1701
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            Height =732
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =465
                    Top =180
                    Width =7755
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="lblComent"
                    Caption ="Transaction Controls && Setup"
                    LeftPadding =30
                    TopPadding =30
                    RightPadding =30
                    BottomPadding =30
                    GridlineStyleLeft =0
                    GridlineStyleTop =0
                    GridlineStyleRight =0
                    GridlineStyleBottom =0
                    GridlineWidthLeft =1
                    GridlineWidthTop =1
                    GridlineWidthRight =1
                    GridlineWidthBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8480
                    Top =188
                    Width =585
                    Name ="yearID"
                    ControlSource ="yearID"

                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5976
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Left =72
                    Top =45
                    Width =9864
                    Height =5835
                    Name ="tabAcc"

                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =408
                            Width =9648
                            Height =5364
                            Name ="Lock Pages"
                            EventProcPrefix ="Lock_Pages"
                            Caption ="Lock Pages"
                            LeftPadding =30
                            TopPadding =30
                            RightPadding =30
                            BottomPadding =30
                            GridlineStyleLeft =0
                            GridlineStyleTop =0
                            GridlineStyleRight =0
                            GridlineStyleBottom =0
                            GridlineWidthLeft =1
                            GridlineWidthTop =1
                            GridlineWidthRight =1
                            GridlineWidthBottom =1
                            Begin
                                Begin OptionGroup
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    Left =330
                                    Top =540
                                    Width =8233
                                    Height =4332
                                    Name ="frmLockPages"
                                    DefaultValue ="1"
                                    OnClick ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin OptionButton
                                            OverlapFlags =215
                                            Left =566
                                            Top =803
                                            OptionValue =1
                                            Name ="optRange"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                            Begin
                                                Begin Label
                                                    OverlapFlags =215
                                                    Left =849
                                                    Top =777
                                                    Width =810
                                                    Height =240
                                                    Name ="Label123"
                                                    Caption ="Ranges"
                                                    LeftPadding =30
                                                    TopPadding =30
                                                    RightPadding =30
                                                    BottomPadding =30
                                                    GridlineStyleLeft =0
                                                    GridlineStyleTop =0
                                                    GridlineStyleRight =0
                                                    GridlineStyleBottom =0
                                                    GridlineWidthLeft =1
                                                    GridlineWidthTop =1
                                                    GridlineWidthRight =1
                                                    GridlineWidthBottom =1
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =215
                                            Left =566
                                            Top =1305
                                            OptionValue =2
                                            Name ="optSelective"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                            Begin
                                                Begin Label
                                                    OverlapFlags =215
                                                    Left =857
                                                    Top =1277
                                                    Width =1530
                                                    Height =240
                                                    Name ="Label125"
                                                    Caption ="Select pages to lock"
                                                    LeftPadding =30
                                                    TopPadding =30
                                                    RightPadding =30
                                                    BottomPadding =30
                                                    GridlineStyleLeft =0
                                                    GridlineStyleTop =0
                                                    GridlineStyleRight =0
                                                    GridlineStyleBottom =0
                                                    GridlineWidthLeft =1
                                                    GridlineWidthTop =1
                                                    GridlineWidthRight =1
                                                    GridlineWidthBottom =1
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =435
                                    Top =660
                                    Width =8005
                                    Height =476
                                    Name ="Box133"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Subform
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =450
                                    Top =1560
                                    Width =6626
                                    Height =3141
                                    TabIndex =1
                                    Name ="frmTransactionLock"
                                    SourceObject ="Form.frmTransactionLock"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3219
                                    Top =787
                                    Width =825
                                    TabIndex =2
                                    Name ="txtFromPage"
                                    Format ="General Number"
                                    ValidationRule =">=0"
                                    ValidationText ="The value must be equal or above 0"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            Left =1755
                                            Top =787
                                            Width =1320
                                            Height =240
                                            Name ="Label137"
                                            Caption ="Page ranges from"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =4586
                                    Top =775
                                    Width =780
                                    TabIndex =3
                                    Name ="txtToPage"
                                    Format ="General Number"
                                    ValidationRule =">=0"
                                    ValidationText ="The value must be equal or above 0"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =4195
                                            Top =794
                                            Width =275
                                            Height =247
                                            Name ="Label143"
                                            Caption ="to"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =6896
                                    Top =780
                                    Width =720
                                    TabIndex =4
                                    Name ="txtUpToPrd"
                                    Format ="General Number"
                                    ValidationRule =">=0"
                                    ValidationText ="The value must be equal or above 0"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5836
                                            Top =780
                                            Width =945
                                            Height =240
                                            Name ="Label141"
                                            Caption ="Up to period"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7680
                                    Top =728
                                    Width =661
                                    Height =330
                                    TabIndex =5
                                    Name ="cmdStartRge"
                                    Caption ="Start"
                                    OnClick ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =7350
                                    Top =4365
                                    Width =969
                                    Height =331
                                    TabIndex =6
                                    Name ="cmdCommit"
                                    Caption ="Commit"
                                    OnClick ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =180
                            Top =408
                            Width =9648
                            Height =5364
                            Name ="pgeInterface"
                            Caption ="Interface"
                            LayoutCachedLeft =135
                            LayoutCachedTop =375
                            LayoutCachedWidth =9630
                            LayoutCachedHeight =5658
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =902
                                    Width =3360
                                    Height =270
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboInvCR"
                                    ControlSource ="ctrInvCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2875
                                            Top =902
                                            Width =1530
                                            Height =255
                                            Name ="lblInvCR"
                                            Caption ="CR Sales A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =2733
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblRcpt"
                                    Caption ="RECEIPTS"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =2686
                                    Width =9403
                                    Height =378
                                    Name ="Box213"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =915
                                    TabIndex =1
                                    Name ="txtInvNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =3191
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblExp"
                                    Caption ="EXPENSES"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =3144
                                    Width =9403
                                    Height =378
                                    Name ="Box217"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =1367
                                    Width =3360
                                    Height =270
                                    TabIndex =2
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboCSH_DR"
                                    ControlSource ="ctrCSH_DR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2875
                                            Top =1367
                                            Width =1530
                                            Height =255
                                            Name ="Label271"
                                            Caption ="DR Sales A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =3649
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblPaymt"
                                    Caption ="PAYMENTS"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =3602
                                    Width =9403
                                    Height =378
                                    Name ="Box221"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =1365
                                    TabIndex =3
                                    Name ="txtCSHNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =4107
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblDepts"
                                    Caption ="DEPOSITS"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =4060
                                    Width =9403
                                    Height =378
                                    Name ="Box225"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =1817
                                    Width =3360
                                    Height =270
                                    TabIndex =4
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboCRN_DR"
                                    ControlSource ="ctrCRN_DR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2875
                                            Top =1817
                                            Width =1530
                                            Height =255
                                            Name ="lblCRN_DR"
                                            Caption ="DR Sales A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =4565
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblWDR"
                                    Caption ="WITHDRAWALS"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =4518
                                    Width =9403
                                    Height =378
                                    Name ="Box229"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =1830
                                    TabIndex =5
                                    Name ="txtCRNNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =2275
                                    Width =3360
                                    Height =270
                                    TabIndex =6
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboDRN_CR"
                                    ControlSource ="ctrDRN_CR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2875
                                            Top =2275
                                            Width =1530
                                            Height =255
                                            Name ="lblDRN_CR"
                                            Caption ="CR Sales A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =1367
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="Label272"
                                    Caption ="CASH SALES"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =300
                                    Top =1320
                                    Width =9403
                                    Height =378
                                    Name ="Box273"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =8100
                                    Top =465
                                    Width =1485
                                    Height =270
                                    Name ="Label274"
                                    Caption ="Document Start No."
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =2280
                                    TabIndex =7
                                    Name ="txtDRNNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =2733
                                    Width =3360
                                    Height =270
                                    TabIndex =8
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboRcptDR"
                                    ControlSource ="ctrRcptDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2875
                                            Top =2733
                                            Width =1530
                                            Height =255
                                            Name ="lblRcptDR"
                                            Caption ="DR Cost A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =2715
                                    TabIndex =9
                                    Name ="txtRecNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =3191
                                    Width =3360
                                    Height =270
                                    TabIndex =10
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboExpDR"
                                    ControlSource ="ctrExpDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2875
                                            Top =3191
                                            Width =1530
                                            Height =255
                                            Name ="lblExpDR"
                                            Caption ="DR Purchase A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =3180
                                    TabIndex =11
                                    Name ="txtExpNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =3649
                                    Width =3360
                                    Height =270
                                    TabIndex =12
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboPaymtCR"
                                    ControlSource ="ctrPaymtCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2773
                                            Top =3645
                                            Width =1635
                                            Height =255
                                            Name ="lblPaymtCR"
                                            Caption ="CR Cash or Bank A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =3630
                                    TabIndex =13
                                    Name ="txtPayNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =4107
                                    Width =3360
                                    Height =270
                                    TabIndex =14
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboDeptsCR"
                                    ControlSource ="ctrDeptsCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName; "
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2875
                                            Top =4107
                                            Width =1530
                                            Height =255
                                            Name ="lblDeptsCR"
                                            Caption ="CR Cash A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =4110
                                    TabIndex =15
                                    Name ="txtDepNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =4565
                                    Width =3360
                                    Height =270
                                    TabIndex =16
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="cboCshWithdrwDR"
                                    ControlSource ="ctrCshWithdrwDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2875
                                            Top =4565
                                            Width =1530
                                            Height =255
                                            Name ="lblCshWithdrwDR"
                                            Caption ="DR Cash A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =902
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblInv"
                                    Caption ="INVOICES "
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =300
                                    Top =855
                                    Width =9403
                                    Height =378
                                    Name ="Box193"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =4560
                                    TabIndex =17
                                    Name ="txtWDRNo"
                                    OnExit ="[Event Procedure]"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =1817
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblCRN"
                                    Caption ="CREDIT NOTES"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =300
                                    Top =1770
                                    Width =9403
                                    Height =378
                                    Name ="Box205"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =5070
                                    Top =500
                                    TabIndex =18
                                    Name ="chkDocAddYear"
                                    ControlSource ="ctrDocAddYear"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5305
                                            Top =465
                                            Width =2535
                                            Height =240
                                            Name ="Label232"
                                            Caption ="Include Year in document numbers"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =2275
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="lblDRN"
                                    Caption ="DEBIT NOTES"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =300
                                    Top =2228
                                    Width =9403
                                    Height =378
                                    Name ="Box209"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =4997
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="Label355"
                                    Caption ="DISCOUNT"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =4950
                                    Width =9403
                                    Height =378
                                    Name ="Box356"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =4997
                                    Width =3360
                                    Height =270
                                    TabIndex =19
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="ctrDiscountDR"
                                    ControlSource ="ctrDiscountDR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2875
                                            Top =4997
                                            Width =1530
                                            Height =255
                                            Name ="Label358"
                                            Caption ="DR Discount A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextFontCharSet =238
                                    Left =388
                                    Top =5417
                                    Width =2355
                                    Height =240
                                    FontWeight =700
                                    Name ="Label360"
                                    Caption ="BANK CHARGES (Inc)"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =300
                                    Top =5370
                                    Width =9403
                                    Height =378
                                    Name ="Box361"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =3402
                                    Left =4494
                                    Top =5417
                                    Width =3360
                                    Height =270
                                    TabIndex =20
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="ctrBankChgCR"
                                    ControlSource ="ctrBankChgCR"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount WH"
                                        "ERE (((tblAccount.accType)=2)) ORDER BY tblAccount.accName;"
                                    ColumnWidths ="0;2268;1134"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2895
                                            Top =5415
                                            Width =1515
                                            Height =255
                                            Name ="Label363"
                                            Caption ="CR Bank Charg A/C"
                                            LeftPadding =30
                                            TopPadding =30
                                            RightPadding =30
                                            BottomPadding =30
                                            GridlineStyleLeft =0
                                            GridlineStyleTop =0
                                            GridlineStyleRight =0
                                            GridlineStyleBottom =0
                                            GridlineWidthLeft =1
                                            GridlineWidthTop =1
                                            GridlineWidthRight =1
                                            GridlineWidthBottom =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8130
                                    Top =5412
                                    TabIndex =21
                                    Name ="Text364"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
                                    ShowDatePicker =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =180
                            Top =408
                            Width =9648
                            Height =5364
                            Name ="pgeCust"
                            Caption ="Customers"
                            LayoutCachedLeft =45
                            LayoutCachedTop =360
                            LayoutCachedWidth =9630
                            LayoutCachedHeight =5655
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =210
                                    Top =450
                                    Width =9435
                                    Height =4800
                                    Name ="frmAccCust"
                                    SourceObject ="Form.frmAccCust"
                                    LeftPadding =30
                                    TopPadding =30
                                    RightPadding =30
                                    BottomPadding =30
                                    GridlineStyleLeft =0
                                    GridlineStyleTop =0
                                    GridlineStyleRight =0
                                    GridlineStyleBottom =0
                                    GridlineWidthLeft =1
                                    GridlineWidthTop =1
                                    GridlineWidthRight =1
                                    GridlineWidthBottom =1
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
                        0x000301000000000000000000
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
                        0x000301000000000000000000
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
                        0x000301000000000000000000
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
                        0x000301000000000000000000
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
                    ControlTipText ="Close Current Form (ALT+X)"

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

Dim sql As String

sql = "UPDATE tblTransaction " & _
      "SET tblTransaction.trnLock = True " & _
      "WHERE tblTransaction.trnPageCounter Between " & Me.txtFromPage & " AND " & Me.txtToPage & _
      " AND tblTransaction.persID <= " & Me.txtUpToPrd
      
DoCmd.SetWarnings False
FnLog (strSQL)
DoCmd.RunSQL sql
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


Private Sub Form_Open(Cancel As Integer)
    txtInvNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'INV' AND cmpYear = " & Me.yearID)
    txtCSHNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'CSH' AND cmpYear = " & Me.yearID)
    txtCRNNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'CRN' AND cmpYear = " & Me.yearID)
    txtDRNNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'DRN' AND cmpYear = " & Me.yearID)
    txtRecNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'REC' AND cmpYear = " & Me.yearID)
    txtExpNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'EXP' AND cmpYear = " & Me.yearID)
    txtPayNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'PAY' AND cmpYear = " & Me.yearID)
    txtDepNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'DEP' AND cmpYear = " & Me.yearID)
    txtWDRNo = DLookup("cmpNo", "UserCmp", "cmpDocType = 'WDR' AND cmpYear = " & Me.yearID)
    
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
    rs.Close
    Set rs = Nothing
End Sub
Private Sub txtInvNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "INV", 1)
End Sub

Private Sub txtCSHNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "CSH", 2)
End Sub

Private Sub txtCRNNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "CRN", 3)
End Sub
Private Sub txtDRNNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "DRN", 4)
End Sub
Private Sub txtRecNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "REC", 5)
End Sub
Private Sub txtExpNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "EXP", 6)
End Sub
Private Sub txtPayNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "PAY", 7)
End Sub

Private Sub txtDepNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "DEP", 8)
End Sub

Private Sub txtWDRNo_Exit(Cancel As Integer)
    Call UDDocNo(Me.ActiveControl, "WDR", 9)
End Sub
Private Sub UDDocNo(lngNewNum As Long, strDocType As String, bytDocType As Byte)

    DoCmd.SetWarnings Flase
    If IsNull(DLookup("cmpNo", "UserCmp", "cmpDocType = '" & strDocType & "' AND cmpYear = " & Me.yearID)) = False Then
        strSQL = "UPDATE UserCmp SET cmpNo = " & lngNewNum & " WHERE cmpDocType = '" & strDocType & "' AND cmpYear = " & Me.yearID
        FnLog (strSQL)
        DoCmd.RunSQL (strSQL)
    Else
        strSQL = "INSERT INTO UserCmp (cmpNo, cmpDocType, cmpDocTypeNo, cmpYear) Values (" & lngNewNum & ",'" & strDocType & "' , " & bytDocType & "," & Me.yearID & ")"
        FnLog (strSQL)
        DoCmd.RunSQL (strSQL)
    End If
    DoCmd.SetWarnings True

End Sub
