Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =238
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6916
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =1065
    Top =1185
    Right =12915
    Bottom =5445
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb6600b2ef87fe340
    End
    RecordSource ="SELECT * FROM tblTransaction WHERE (((tblTransaction.trnLock)=0)) ORDER BY tblTr"
        "ansaction.trnPageCounter; "
    Caption ="tblTransaction"
    DatasheetFontName ="Arial"
    Moveable =0
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =238
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            TextFontCharSet =238
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =238
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =238
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =238
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            BackStyle =0
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =354
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =57
                    Width =1035
                    Height =240
                    Name ="lblEntryDate"
                    Caption ="Entry Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1152
                    Top =57
                    Width =1185
                    Height =240
                    Name ="trnInternalRef_Label"
                    Caption ="Voucher"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2394
                    Top =57
                    Width =1185
                    Height =240
                    Name ="trnPageCounter_Label"
                    Caption ="Page Number"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3636
                    Top =57
                    Width =600
                    Height =240
                    Name ="persID_Label"
                    Caption ="Period"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4293
                    Top =57
                    Width =900
                    Height =240
                    Name ="trnYear_Label"
                    Caption ="Year"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5250
                    Top =57
                    Width =615
                    Height =240
                    Name ="trnLock_Label"
                    Caption ="Lock"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =369
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =57
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="trnEntryDate"
                    ControlSource ="trnEntryDate"
                    Format ="Short Date"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1152
                    Top =57
                    Width =1185
                    Height =255
                    ColumnWidth =1185
                    TabIndex =1
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2394
                    Top =57
                    Width =1185
                    Height =255
                    ColumnWidth =900
                    TabIndex =2
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="trnPageCounter"
                    ControlSource ="trnPageCounter"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3636
                    Top =57
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="persID"
                    ControlSource ="persID"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4293
                    Top =57
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =4
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="trnYear"
                    ControlSource ="trnYear"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5400
                    Top =120
                    TabIndex =5
                    Name ="trnLock"
                    ControlSource ="trnLock"
                    DefaultValue ="0"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
