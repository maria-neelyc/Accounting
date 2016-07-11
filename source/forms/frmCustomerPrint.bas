Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3401
    DatasheetFontHeight =10
    ItemSuffix =77
    Left =4215
    Top =60
    Right =7620
    Bottom =2760
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x675d19c84f96e140
    End
    Caption ="Print Customer List"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6e0400006e0400006e0400006e04000000000000041b0000660c000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =2721
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =80
                    Left =1303
                    Top =1303
                    Width =1206
                    Height =456
                    TabIndex =1
                    Name ="cmdPrint"
                    Caption ="&Print"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad00000000000dadd0888888888080da ,
                        0x000000000000080d0888888bbb88000a088888877788080d0000000000000880 ,
                        0x0888888888808080d000000000080800ad0ffffffff08080dad0f00000f0000a ,
                        0xada0ffffffff0daddada0f00000f0adaadad0ffffffff0addadad000000000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Print"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =1303
                    Top =1813
                    Width =1206
                    Height =456
                    TabIndex =2
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Exit"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =1307
                    Top =176
                    Width =1187
                    Height =943
                    Name ="frmPrintOn"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =1427
                            Top =56
                            Width =720
                            Height =240
                            FontWeight =700
                            BackColor =-2147483633
                            Name ="Label27"
                            Caption ="Print on"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =1493
                            Top =414
                            OptionValue =1
                            Name ="Check29"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1723
                                    Top =386
                                    Width =585
                                    Height =240
                                    Name ="Label30"
                                    Caption ="Screen"
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =1493
                            Top =744
                            OptionValue =2
                            Name ="Check31"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1723
                                    Top =716
                                    Width =525
                                    Height =240
                                    Name ="Label32"
                                    Caption ="Printer"
                                End
                            End
                        End
                    End
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
Option Explicit
Private Sub cmdExit_Click()
DoCmd.Close
End Sub
Private Sub cmdPrint_Click()
Dim strReport As String
On Error GoTo Err_End


strReport = "rptCustomer"

Select Case frmPrintOn
Case 1
    DoCmd.OpenReport strReport, acViewPreview
Case 2
    DoCmd.OpenReport strReport, acViewNormal
End Select

Err_End:
End Sub



Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim strMsg As String
       strMsg = FnIsErr(DataErr)
       Response = acDataErrContinue
       If strMsg <> "" Or IsNull(strMsg) Or IsEmpty(strMsg) Then
           MsgBox strMsg, vbInformation, "Error!"
       Else
       End If
End Sub
