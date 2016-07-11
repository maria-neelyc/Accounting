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
    Width =8280
    DatasheetFontHeight =10
    ItemSuffix =45
    Left =705
    Top =255
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd3aef4b5189be240
    End
    RecordSource ="Order Details with ShippedDate"
    Caption ="Sales by product"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a00500000000000058200000f000000001000000 ,
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="ProductName"
        End
        Begin BreakLevel
            ControlSource ="OrderID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1200
            OnPrint ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =780
                    Width =1440
                    Height =300
                    Name ="ProductName_Label"
                    Caption ="Product Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =1620
                    Top =780
                    Width =1020
                    Height =300
                    Name ="OrderID_Label"
                    Caption ="Order ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =2700
                    Top =780
                    Width =1140
                    Height =300
                    Name ="ShippedDate_Label"
                    Caption ="Shipped on"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    Left =3900
                    Top =780
                    Width =2040
                    Height =300
                    Name ="CompanyName_Label"
                    Caption ="Company"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =6000
                    Top =780
                    Width =900
                    Height =300
                    Name ="ExtendedPrice_Label"
                    Caption ="Sale"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =7020
                    Top =780
                    Width =1200
                    Height =300
                    Name ="ProductID_Label"
                    Caption ="Cumulative"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =8280
                    Name ="Line23"
                End
                Begin Line
                    BorderWidth =2
                    Top =1140
                    Width =8280
                    Name ="Line24"
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    Top =60
                    Width =3240
                    Height =510
                    FontSize =20
                    Name ="Label19"
                    Caption ="Sales by product"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =8280
                    BorderColor =12632256
                    Name ="Line22"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =330
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =2460
                    Height =330
                    ColumnWidth =3210
                    FontSize =11
                    ForeColor =8388608
                    Name ="ProductName"
                    ControlSource ="ProductName"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =1860
                    Width =780
                    Height =225
                    ColumnWidth =945
                    Name ="OrderID"
                    ControlSource ="OrderID"
                    StatusBarText ="Same as Order ID in Orders table."

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2700
                    Width =1140
                    Height =225
                    ColumnWidth =1440
                    TabIndex =1
                    Name ="ShippedDate"
                    ControlSource ="ShippedDate"
                    Format ="dd-mmm-yyyy"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3900
                    Width =2040
                    Height =225
                    ColumnWidth =3390
                    TabIndex =2
                    Name ="CompanyName"
                    ControlSource ="CompanyName"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6000
                    Width =1020
                    Height =225
                    ColumnWidth =1545
                    TabIndex =3
                    Name ="ExtendedPrice"
                    ControlSource ="ExtendedPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin TextBox
                    RunningSum =1
                    IMESentenceMode =3
                    Left =7260
                    Width =1020
                    Height =225
                    TabIndex =4
                    Name ="Cumulative"
                    ControlSource ="ExtendedPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =840
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =4980
                    FontSize =11
                    ForeColor =8388608
                    Name ="Text12"
                    ControlSource ="=\"Total for \" & [ProductName] & \" (\" & Count(*) & \" \" & IIf(Count(*)=1,\"d"
                        "etail record\",\"orders\") & \")\""
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =0
                    Left =60
                    Top =300
                    Width =885
                    Height =270
                    FontWeight =400
                    Name ="Label14"
                    Caption ="Percent"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5220
                    Top =60
                    Width =1800
                    Height =270
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Sum Of ExtendedPrice"
                    ControlSource ="=Sum([ExtendedPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Sum_Of_ExtendedPrice"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5100
                    Top =360
                    Width =1800
                    Height =270
                    FontSize =11
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Standard Of ExtendedPrice"
                    ControlSource ="=Sum([ExtendedPrice])/([ExtendedPrice Grand Total Sum])"
                    Format ="Percent"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Standard_Of_ExtendedPrice"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Top =720
                    Width =8220
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    Left =5760
                    Width =1260
                    Name ="Line29"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    RunningSum =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7020
                    Top =60
                    Width =1260
                    Height =270
                    TabIndex =3
                    ForeColor =8388608
                    Name ="txtRunningGrpSum"
                    ControlSource ="=Sum([ExtendedPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"

                End
                Begin Line
                    Left =7140
                    Width =1140
                    Name ="Line44"
                End
            End
        End
        Begin PageFooter
            Height =780
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =3900
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8421504
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
                    Left =4500
                    Top =420
                    Width =3720
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8421504
                    Name ="Text21"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5220
                    Top =180
                    Width =1800
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8421504
                    Name ="txtPageSum"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =3
                    Left =3540
                    Top =180
                    Width =1680
                    Height =300
                    FontSize =9
                    ForeColor =8421504
                    Name ="Label39"
                    Caption ="Total for this page:"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =180
                    Width =8280
                    BorderColor =12632256
                    Name ="Line41"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =480
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5580
                    Top =60
                    Height =285
                    FontSize =11
                    Name ="ExtendedPrice Grand Total Sum"
                    ControlSource ="=Sum([ExtendedPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"
                    EventProcPrefix ="ExtendedPrice_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =60
                    Width =1200
                    Height =300
                    ForeColor =0
                    Name ="Label16"
                    Caption ="Grand Total"
                End
                Begin Line
                    Left =5616
                    Top =360
                    Width =1404
                    Name ="Line31"
                End
                Begin Line
                    Left =5616
                    Top =420
                    Width =1404
                    Name ="Line32"
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

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
If PrintCount = 1 Then
   txtPageSum = txtPageSum + ExtendedPrice
End If
End Sub

Private Sub PageHeaderSection_Print(Cancel As Integer, PrintCount As Integer)
txtPageSum = 0
End Sub
