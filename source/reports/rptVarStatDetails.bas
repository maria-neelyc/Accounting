Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11111
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =135
    Top =75
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x00545d70b79ce340
    End
    RecordSource ="qryVatStatDetails"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xc0000000a8000000c00000001e01000000000000672b00005401000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            TextFontFamily =2
            FontName ="Arial"
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
            Width =850
            Height =850
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
        Begin Line
            Width =1701
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
            OldBorderStyle =0
            TextFontCharSet =238
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
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
            ShowDatePicker =0
        End
        Begin Section
            KeepTogether = NotDefault
            Height =340
            BackColor =12632256
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1485
                    Top =60
                    Width =906
                    BackColor =12632256
                    Name ="trnEntryDate"
                    ControlSource ="trnEntryDate"
                    Format ="Short Date"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2430
                    Top =60
                    Width =1071
                    TabIndex =1
                    BackColor =12632256
                    Name ="trnsNote"
                    ControlSource ="refName"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3525
                    Top =53
                    Width =501
                    TabIndex =2
                    BackColor =12632256
                    Name ="vatName"
                    ControlSource ="vatName"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =56
                    Width =1356
                    TabIndex =3
                    BackColor =12632256
                    Name ="trnInternalRef"
                    ControlSource ="trnInternalRef"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =56
                    Width =2226
                    TabIndex =4
                    BackColor =12632256
                    Name ="desName"
                    ControlSource ="trnsNote"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Top =60
                    Width =1134
                    TabIndex =5
                    BackColor =12632256
                    Name ="txttrnsDebits"
                    ControlSource ="=IIf(IsNumeric([trnsDebits]),IIf([trnsDebits]=0,\"\",[trnsDebits]),\"\")"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8385
                    Top =60
                    Width =1134
                    TabIndex =6
                    BackColor =12632256
                    Name ="txttrnsCredits"
                    ControlSource ="=IIf(IsNumeric([trnsCredits]),IIf([trnsCredits]=0,\"\",[trnsCredits]),\"\")"
                    Format ="Standard"

                End
                Begin Line
                    Left =8392
                    Width =0
                    Height =340
                    Name ="Line11"
                End
                Begin Line
                    Left =1424
                    Width =0
                    Height =340
                    Name ="Line13"
                End
                Begin Line
                    Left =2385
                    Width =0
                    Height =340
                    Name ="Line14"
                End
                Begin Line
                    Left =4773
                    Width =0
                    Height =340
                    Name ="Line16"
                End
                Begin Line
                    Left =7230
                    Width =0
                    Height =340
                    Name ="Line17"
                End
                Begin Line
                    Left =3515
                    Width =0
                    Height =340
                    Name ="Line30"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4185
                    Top =53
                    Width =501
                    TabIndex =7
                    BackColor =12632256
                    Name ="vatRate"
                    ControlSource ="vatRate"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4027
                    Top =46
                    Width =126
                    TabIndex =8
                    BackColor =12632256
                    Name ="txtDash"
                    ControlSource ="=\"-\""

                End
                Begin Line
                    Left =9537
                    Width =0
                    Height =340
                    Name ="Line3"
                End
            End
        End
    End
End
