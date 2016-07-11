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
    Width =9411
    DatasheetFontHeight =10
    ItemSuffix =53
    Left =495
    Top =1575
    Right =13200
    Bottom =5865
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa73bc5deebc5e440
    End
    RecordSource ="SELECT appCustomer.custID, appCustomer.custSName, appCustomer.custName, appCusto"
        "mer.custAddress, appCustomer.custTown, appCustomer.custBankAccNo, appCustomer.ac"
        "cID FROM appCustomer WHERE appCustomer.accID = 0"
    Caption ="tblTransaction"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xae050000ae050000ae050000ae05000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Moveable =0
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
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
        Begin CustomControl
            SpecialEffect =2
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
            Height =1077
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin OptionGroup
                    OverlapFlags =85
                    Left =225
                    Top =245
                    Width =3497
                    Height =688
                    ColumnOrder =0
                    Name ="grpAccCust"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =350
                            Top =120
                            Width =1575
                            Height =240
                            Name ="Label38"
                            Caption ="Show only customers"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =411
                            Top =483
                            OptionValue =1
                            Name ="optAccCust1"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =641
                                    Top =455
                                    Width =1350
                                    Height =240
                                    Name ="Label41"
                                    Caption ="Without Accounts"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2210
                            Top =493
                            OptionValue =2
                            Name ="optAccCust2"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2440
                                    Top =465
                                    Width =1020
                                    Height =240
                                    Name ="Label43"
                                    Caption ="All Customers"
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =4035
                    Top =245
                    Width =3497
                    Height =688
                    ColumnOrder =1
                    TabIndex =1
                    Name ="grpAccID"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =4160
                            Top =120
                            Width =1245
                            Height =240
                            Name ="Label45"
                            Caption ="Show Accounts"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =4221
                            Top =483
                            OptionValue =1
                            Name ="optAccID1"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =4451
                                    Top =455
                                    Width =1350
                                    Height =240
                                    Name ="Label47"
                                    Caption ="Available"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6020
                            Top =493
                            OptionValue =2
                            Name ="optAccID2"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6250
                                    Top =465
                                    Width =1020
                                    Height =240
                                    Name ="Label49"
                                    Caption ="All"
                                End
                            End
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    SpecialEffect =1
                    Top =1020
                    Width =9081
                    Name ="Line50"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =7766
                    Top =283
                    Width =1311
                    Height =576
                    TabIndex =2
                    Name ="cmdCrCust"
                    Caption ="&Create Customers"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Create Customers  (ALT+C)"
                    UnicodeAccessKey =67

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =1587
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3345
                    Top =225
                    Width =735
                    Height =255
                    ColumnWidth =1035
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="custID"
                    ControlSource ="custID"
                    Format ="Short Date"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =277
                    Top =509
                    Width =3930
                    Height =255
                    ColumnWidth =1185
                    TabIndex =1
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="custName"
                    ControlSource ="custName"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4282
                    Top =509
                    Width =1800
                    Height =855
                    ColumnWidth =900
                    TabIndex =2
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="trnPageCounter"
                    ControlSource ="custAddress"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2669
                    Top =1076
                    Width =1545
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="persID"
                    ControlSource ="custTown"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =277
                    Top =1076
                    Width =2340
                    Height =255
                    ColumnWidth =900
                    TabIndex =4
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="custBankAccNo"
                    ControlSource ="custBankAccNo"

                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =277
                    Top =803
                    Width =2340
                    Height =240
                    Name ="lblEntryDate"
                    Caption ="Bank Account"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2669
                    Top =803
                    Width =1576
                    Height =240
                    Name ="trnInternalRef_Label"
                    Caption ="City"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =277
                    Top =225
                    Width =1215
                    Height =240
                    Name ="trnPageCounter_Label"
                    Caption ="Cusomer"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =4297
                    Top =229
                    Width =750
                    Height =240
                    Name ="persID_Label"
                    Caption ="Address"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =7995
                    Top =165
                    Width =915
                    Height =240
                    Name ="trnYear_Label"
                    Caption ="Account"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3402
                    Left =6525
                    Top =450
                    Width =2382
                    Height =284
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"24\""
                    Name ="cboCustID"
                    ControlSource ="accID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName FROM appCustomer R"
                        "IGHT JOIN tblAccount ON appCustomer.accID = tblAccount.accID WHERE ((([tblAccoun"
                        "t].[accID] & [appCustomer].[accID])=[tblAccount].[accID])) ORDER BY tblAccount.a"
                        "ccName; "
                    ColumnWidths ="0;1134;2268"
                    OnExit ="[Event Procedure]"
                    SpecialEffect =1
                    OverlapFlags =255
                    Left =165
                    Top =165
                    Width =6183
                    Height =1251
                    Name ="Box22"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =113
                    Top =113
                    Width =8901
                    Height =1361
                    Name ="Box23"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1545
                    Top =225
                    Width =1635
                    Height =255
                    TabIndex =6
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =8421504
                    Name ="custSName"
                    ControlSource ="custSName"
                    Format ="Short Date"

                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =65
                    TextFontCharSet =0
                    Left =7935
                    Top =1020
                    Width =967
                    Height =360
                    TabIndex =7
                    Name ="cmdAdd"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Account  (ALT+A)"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboCustID_Exit(Cancel As Integer)
    Dim bytNoOthAcc As Byte
    
    'bytNoOthAcc = DLookup("Count(accID)", "appCustomer", "accID = accID" & cboCustID)
'    If bytNoOthAcc > 1 Then
'        lblOthAcc.Caption = bytNoOthAcc
'        Me.lblOthAcc.ForeColor = Red
'    Else
'        Me.lblOthAcc.ForeColor = 0
'    End If
    
    Dim rst As Recordset
    Dim strSQL As String

'    strSQL = "SELECT tblAccount.accNo, tblAccount.accName FROM tblAccount WHERE tblAccount.caoID = " & caoID & " AND tblAccount.accNo <> '" & accNo & "'"
    If cboCustID <> 0 Then
        strSQL = "SELECT appCustomer.custID, appCustomer.custSName, appCustomer.custName FROM appCustomer WHERE appCustomer.accID = " & cboCustID & " AND appCustomer.custID <> " & custID

        Set rst = CurrentDb.OpenRecordset(strSQL)
        If rst.RecordCount > 0 Then
            MsgBox ("This reference is already assigned to customer: " & rst.Fields(1) & " - " & rst.Fields(2))
            DoCmd.CancelEvent
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    DoCmd.OpenForm "frmAccount", , , , , , Me.Form.name
    'Me.cboCustID.Requery
    Me.cboCustID.Requery
    Me.Refresh
End Sub
Private Sub cmdCrCust_Click()
    DoCmd.OpenForm "frmCrCust"
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "SELECT appCustomer.custID, appCustomer.custSName, appCustomer.custName, appCustomer.custAddress, appCustomer.custTown, appCustomer.custBankAccNo, appCustomer.accID " & _
                                         "FROM appCustomer " & _
                                         "WHERE appCustomer.accID = 0"
    Me.cboCustID.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
                                     "FROM appCustomer RIGHT JOIN tblAccount ON appCustomer.accID = tblAccount.accID " & _
                                     "WHERE ((([tblAccount].[accID] & [appCustomer].[accID])=[tblAccount].[accID])) " & _
                                     "ORDER BY tblAccount.accName; "
    Me.Requery
    Me.cboCustID.Requery
End Sub

Private Sub grpAccCust_Click()
    
    Select Case grpAccCust
        Case 1:
            Me.RecordSource = "SELECT appCustomer.custID, appCustomer.custSName, appCustomer.custName, appCustomer.custAddress, appCustomer.custTown, appCustomer.custBankAccNo, appCustomer.accID " & _
                                         "FROM appCustomer " & _
                                         "WHERE appCustomer.accID = 0"
'            Call grpAccID_Click
'            grpAccID.Enabled = True
        Case 2:
            Me.RecordSource = "SELECT appCustomer.custID, appCustomer.custSName, appCustomer.custName, appCustomer.custAddress, appCustomer.custTown, appCustomer.custBankAccNo, appCustomer.accID " & _
                                         "FROM appCustomer "
            Me.cboCustID.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
                                     "FROM appCustomer RIGHT JOIN tblAccount ON appCustomer.accID = tblAccount.accID " & _
                                     "ORDER BY tblAccount.accName; "
            grpAccID = 2
'            Me.cboCustID.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
'                                     "FROM appCustomer RIGHT JOIN tblAccount ON appCustomer.accID = tblAccount.accID "
'            grpAccID.Enabled = False
    End Select
    
    Me.Requery
    Me.cboCustID.Requery
End Sub

Private Sub grpAccID_Click()
    
    Select Case grpAccID
        Case 1:
            Me.cboCustID.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
                                     "FROM appCustomer RIGHT JOIN tblAccount ON appCustomer.accID = tblAccount.accID " & _
                                     "WHERE ((([tblAccount].[accID] & [appCustomer].[accID])=[tblAccount].[accID])) " & _
                                     "ORDER BY tblAccount.accName; "
        Case 2:
            Me.cboCustID.RowSource = "SELECT tblAccount.accID, tblAccount.accNo, tblAccount.accName " & _
                                     "FROM appCustomer RIGHT JOIN tblAccount ON appCustomer.accID = tblAccount.accID " & _
                                     "ORDER BY tblAccount.accName; "
    End Select
    
    Me.Requery
    Me.cboCustID.Requery

End Sub
Private Sub Command52_Click()
On Error GoTo Err_Command52_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command52_Click:
    Exit Sub

Err_Command52_Click:
    MsgBox Err.Description
    Resume Exit_Command52_Click
    
End Sub
