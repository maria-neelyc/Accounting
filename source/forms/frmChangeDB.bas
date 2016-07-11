Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6968
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =5400
    Top =795
    Right =12375
    Bottom =7725
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xec2d476c50dae240
    End
    Caption ="EasyBook - Choose Company"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =6944
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =2694
                    Top =330
                    Width =3630
                    Height =510
                    FontSize =20
                    ForeColor =10040115
                    Name ="Label10"
                    Caption ="Ledger  EasyBook"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =3110
                    Top =900
                    Width =3180
                    Height =210
                    FontWeight =700
                    ForeColor =10040115
                    Name ="Label11"
                    Caption ="Easy Bookeeping System"
                End
                Begin Rectangle
                    SpecialEffect =5
                    BackStyle =1
                    OverlapFlags =87
                    Left =354
                    Top =840
                    Width =6000
                    Height =60
                    BackColor =10040115
                    Name ="Box12"
                End
                Begin Label
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Top =6600
                    Width =6825
                    Height =255
                    Name ="Label24"
                    Caption ="Copyright (c) Neelyc Software  LTD 2009   All rights reserved."
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4835
                    Top =1140
                    Width =1455
                    Height =210
                    ForeColor =10040115
                    Name ="Label25"
                    Caption ="v.2009.11.12"
                End
                Begin ListBox
                    OverlapFlags =85
                    ColumnCount =5
                    Left =590
                    Top =2013
                    Width =5625
                    Height =3354
                    Name ="lstMyCo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT sysMyCo.smycoID, sysMyCo.smycoName, sysMyCo.smycoPath, sysMyCo.smycoYear,"
                        " sysMyCo.smycoSubDB1 FROM sysMyCo ORDER BY sysMyCo.smycoID, sysMyCo.smycoName, s"
                        "ysMyCo.smycoYear DESC;"
                    ColumnWidths ="0;5103;0;567;0"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =566
                            Top =1700
                            Width =2535
                            Height =240
                            FontWeight =700
                            Name ="Label27"
                            Caption ="Choose Company:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =590
                    Top =5480
                    Width =5605
                    Height =360
                    TabIndex =1
                    Name ="cmdSelect"
                    Caption ="&Select"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Load selected database (ALT+S)"
                    UnicodeAccessKey =83

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =66
                    Left =377
                    Top =425
                    Width =1525
                    Height =360
                    TabIndex =2
                    Name ="cmdBack"
                    Caption ="&Back"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Go to main screen"
                    UnicodeAccessKey =66

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

Private Sub cmdBack_Click()
    DoCmd.OpenForm "frmMain"
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdSelect_Click()
    
    If IsNoData(lstMyCo.Column(2)) = False Then
        
        If fRefreshLinks(lstMyCo.Column(2), False) = True Then
            Me.Visible = False
            DoCmd.SetWarnings False
            'DoCmd.RunSQL "UPDATE tblControl SET ctrDBID = " & lstMyCo.Column(0)
             strSQL = "UPDATE tblControl SET tblControl.ctrName = ' " & lstMyCo.Column(1) & "' , tblControl.yearID =  " & lstMyCo.Column(3) & ", tblControl.ctrdbID =  " & lstMyCo.Column(0)
             FnLog (strSQL)
             DoCmd.RunSQL (strSQL)
            'UPDATE tblControl SET tblControl.ctrName = [sysMyCo].[smyconame], tblControl.yearID = [sysMyco].[smycoyear];

            DoCmd.SetWarnings True
            DoCmd.OpenForm "frmMain"
            DoCmd.Close acForm, "frmStartUp"
            Exit Sub
        Else
        
        End If
    Else
        MsgBox "Please select Company from List.", vbInformation
    End If
End Sub

Private Sub lstMyCo_DblClick(Cancel As Integer)
    Call cmdSelect_Click
End Sub
