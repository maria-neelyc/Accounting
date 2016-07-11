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
    Width =9259
    DatasheetFontHeight =10
    ItemSuffix =51
    Left =1530
    Top =285
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x926524bcf46ee340
    End
    RecordSource ="SELECT tblChart.coaID, tblChart.coaIDg, tblChart.coaNo, tblChart.coaName FROM tb"
        "lChart; "
    Caption ="Sales by product"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6e0400006e0400005303000053030000000000002b240000f000000001000000 ,
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =4086
                    Name ="coaNo"
                    ControlSource ="coaNo"

                End
                Begin TextBox
                    Left =5858
                    Width =3180
                    ColumnWidth =1545
                    TabIndex =1
                    Name ="coaName"
                    ControlSource ="coaName"

                End
                Begin TextBox
                    TabIndex =2
                    Name ="coaID"
                    ControlSource ="coaID"

                End
                Begin TextBox
                    Left =1606
                    Width =2385
                    TabIndex =3
                    Name ="coaIDg"
                    ControlSource ="coaIDg"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
