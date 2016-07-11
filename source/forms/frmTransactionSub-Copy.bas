Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9977
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =525
    Top =915
    Right =10500
    Bottom =6075
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6fbd4448d2dce240
    End
    RecordSource ="SELECT tblTransactionSub.* FROM tblTransactionSub; "
    DatasheetFontName ="Arial"
    FilterOnLoad =0
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
        Begin FormHeader
            Height =420
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =195
                    Top =180
                    Width =765
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
                    Left =2462
                    Top =180
                    Width =720
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
                    Left =3573
                    Top =180
                    Width =555
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label94"
                    Caption ="Curr."
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =4423
                    Top =180
                    Width =405
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label95"
                    Caption ="Ref."
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =5203
                    Top =150
                    Width =615
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label96"
                    Caption ="Desc."
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =6668
                    Top =180
                    Width =765
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label97"
                    Caption ="F.Amount"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =8085
                    Top =180
                    Width =765
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label98"
                    Caption ="Amount"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =6029
                    Top =180
                    Width =525
                    Height =210
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label99"
                    Caption ="Rate"
                    FontName ="MS Sans Serif"
                End
            End
        End
        Begin Section
            Height =340
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2462
                    Top =67
                    Width =975
                    Height =255
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsDate"
                    ControlSource ="trnsDate"
                    Format ="Short Date"
                    FontName ="MS Sans Serif"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3119
                    Left =165
                    Top =52
                    Width =2250
                    Height =270
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="accID"
                    ControlSource ="accID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accName, tblAccount.accNo FROM tblAccount OR"
                        "DER BY tblAccount.accName; "
                    ColumnWidths ="0;2268;852"
                    FontName ="MS Sans Serif"
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =1701
                    Left =3596
                    Top =52
                    Width =690
                    Height =270
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="crnID"
                    ControlSource ="crnID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCurrency.crnID, tblCurrency.crnShortName, tblCurrency.crnDefault, tblC"
                        "urrency.crnExchangeRate FROM tblCurrency ORDER BY tblCurrency.crnDefault, tblCur"
                        "rency.crnShortName; "
                    ColumnWidths ="0;567;0;1134"
                    FontName ="MS Sans Serif"
                    OnChange ="[Event Procedure]"
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =4399
                    Top =52
                    Width =750
                    Height =270
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="refID"
                    ControlSource ="refID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblReference.refID, tblReference.refShortName, tblReference.refName FROM "
                        "tblReference ORDER BY tblReference.refShortName; "
                    ColumnWidths ="0;567;1701"
                    FontName ="MS Sans Serif"
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2268
                    Left =5203
                    Top =52
                    Width =750
                    Height =270
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="desID"
                    ControlSource ="desID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblDescription.desID, tblDescription.desShortName, tblDescription.desName"
                        " FROM tblDescription ORDER BY tblDescription.desShortName; "
                    ColumnWidths ="0;567;1701"
                    FontName ="MS Sans Serif"
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6644
                    Top =67
                    Width =1365
                    Height =255
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsFAmount"
                    ControlSource ="trnsFAmount"
                    Format ="Standard"
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
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8062
                    Top =67
                    Width =1365
                    Height =255
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsAmount"
                    ControlSource ="trnsAmount"
                    Format ="Standard"
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6029
                    Top =67
                    Width =540
                    Height =255
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsRate"
                    ControlSource ="trnsRate"
                    Format ="Standard"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8510
                    Top =67
                    Width =210
                    Height =255
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsVAmount"
                    ControlSource ="trnsVAmount"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    Left =835
                    Top =30
                    Width =255
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnID"
                    ControlSource ="trnID"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    Left =277
                    Top =67
                    Width =480
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="trnsID"
                    ControlSource ="trnsID"
                    FontName ="MS Sans Serif"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9495
                    Top =45
                    Width =381
                    Height =291
                    TabIndex =11
                    ForeColor =0
                    Name ="Command15"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadad00000000dadadada ,
                        0xfffffff0adadadadfffffff0dadadadafffffff0adadadad00000000000ad4da ,
                        0xefefefefef0da44dfefefefefe044444efefefefef0da44d00000000000ad4da ,
                        0xfffffff0adadadadfffffff0dadadadafffffff0adadadad00000000dadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record (Alt+N)"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =859
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =5
                    Left =195
                    Top =60
                    Width =9258
                    Name ="Line100"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6576
                    Top =113
                    Width =1365
                    Height =225
                    BackColor =-2147483643
                    ForeColor =32768
                    Name ="Text11"
                    ControlSource ="=Sum([trnsFamount])"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000004000000000000000200000001000000 ,
                        0x0000ff00ffffff00000000000500000003000000050000000100000080000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100020000000000000004000000010000000000ff00ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0080000000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7994
                    Top =113
                    Width =1365
                    Height =225
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =32768
                    Name ="Text12"
                    ControlSource ="=Sum([trnsAmount])"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000004000000000000000200000001000000 ,
                        0x0000ff00ffffff00000000000500000003000000050000000100000080000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100020000000000000004000000010000000000ff00ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0080000000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =4935
                    Top =120
                    Width =1575
                    Height =225
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label13"
                    Caption ="Controm Amount:"
                    FontName ="MS Sans Serif"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6591
                    Top =453
                    Width =2820
                    Height =302
                    TabIndex =2
                    Name ="Command17"
                    Caption ="Add New Transaction"

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

Private Sub crnID_Change()
    trnsRate.Value = crnID.Column(3)
    If crnID.Column(2) = 0 Then
        trnsFAmount.Enabled = True
    Else
        trnsFAmount.Value = 0
        trnsFAmount.Enabled = False
    End If
End Sub
