Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10155
    DatasheetFontHeight =10
    ItemSuffix =68
    Left =1935
    Top =285
    DatasheetGridlinesColor =12632256
    Toolbar ="PrintReportEmail"
    RecSrcDt = Begin
        0xfe4a85c46be1e240
    End
    RecordSource ="tblCurrency"
    Caption ="Sales by product"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6e04000037020000370200003702000000000000ab2700000e01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            ControlSource ="crnDefault"
        End
        Begin BreakLevel
            ControlSource ="crnShortName"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1777
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =570
                    Top =1410
                    Width =1440
                    Height =300
                    ForeColor =0
                    Name ="ProductName_Label"
                    Caption ="Short Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =2187
                    Top =1410
                    Width =2895
                    Height =300
                    ForeColor =0
                    Name ="OrderID_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10155
                    BorderColor =0
                    Name ="Line23"
                End
                Begin Line
                    BorderWidth =2
                    Top =1777
                    Width =10140
                    BorderColor =0
                    Name ="Line24"
                End
                Begin Label
                    BackStyle =1
                    Left =70
                    Top =614
                    Width =4470
                    Height =510
                    FontSize =20
                    ForeColor =0
                    Name ="Label19"
                    Caption ="Currency List"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =18
                    Top =386
                    Width =10125
                    Height =15
                    BorderColor =0
                    Name ="Line22"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    Left =540
                    Top =45
                    Width =8565
                    Height =255
                    FontSize =9
                    ForeColor =0
                    Name ="Text37"
                    Caption ="DARE TRADING LTD - Ladger Easy Book"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =2187
                    Top =1410
                    Width =2895
                    Height =300
                    ForeColor =0
                    Name ="Label62"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =5669
                    Top =1440
                    Width =1695
                    Height =300
                    ForeColor =0
                    Name ="Label65"
                    Caption ="Rate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    Left =7653
                    Top =1440
                    Width =765
                    Height =300
                    ForeColor =0
                    Name ="Label67"
                    Caption ="Base"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =270
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =566
                    Width =1425
                    FontSize =10
                    Name ="coaNo3"
                    ControlSource ="crnShortName"

                End
                Begin TextBox
                    Left =2243
                    Width =2865
                    ColumnWidth =6315
                    FontSize =10
                    TabIndex =1
                    Name ="coaName3"
                    ControlSource ="crnName"

                End
                Begin TextBox
                    TextAlign =1
                    Left =5692
                    Width =1635
                    FontSize =10
                    TabIndex =2
                    Name ="Text64"
                    ControlSource ="crnExchangeRate"

                End
                Begin CheckBox
                    Left =7998
                    Top =30
                    TabIndex =3
                    Name ="crnDefault"
                    ControlSource ="crnDefault"

                End
            End
        End
        Begin PageFooter
            Height =720
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =30
                    Top =45
                    Width =2775
                    Height =300
                    FontSize =9
                    FontWeight =700
                    Name ="Text20"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7290
                    Top =60
                    Width =2835
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Text21"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin Line
                    BorderWidth =2
                    Width =10140
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Label
                    TextAlign =2
                    Left =1140
                    Top =450
                    Width =7380
                    Height =270
                    FontSize =8
                    ForeColor =0
                    Name ="Label38"
                    Caption ="Copyright (c) DARE TRADING LTD."
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
