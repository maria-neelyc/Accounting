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
    Width =8340
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =3480
    Top =4095
    Right =12885
    Bottom =12225
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb95ed4bbf46ee340
    End
    RecordSource ="SELECT tblBankAccount.* FROM tblBankAccount WHERE (((tblBankAccount.cstID)=1)); "
    Caption ="Bank Accounts"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnError ="[Event Procedure]"
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
            BackColor =-2147483633
            ForeColor =-2147483630
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
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1260
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Image
                    Left =300
                    Top =180
                    Width =900
                    Height =1020
                    Name ="imgHeader"
                    PictureData = Begin
                        0x030000000000000008000000c905000094050000000000000100090000032e0c ,
                        0x00000300c603000000000400000003010800050000000b020000000005000000 ,
                        0x0c02360038000500000006010200000005000000020101000000050000000102 ,
                        0xb7c94c0007000000fc020000b7c94c000000040000002d01000008000000fa02 ,
                        0x05000100000000000000040000002d010100c60300002403e1010f0000000f00 ,
                        0x00000f0000000f0000000f000000100002001100030012000400130004001400 ,
                        0x05001500060016000700170007001800080019000900190009001a000a001b00 ,
                        0x0a001c000a001d000b001e000b001d000b001c000b001b000a001a000a001900 ,
                        0x0a00180009001700090016000900150008001400080013000700120006001100 ,
                        0x0600100005000f0004000e0004000e0004000e0004000e0004000e0004000f00 ,
                        0x0500100006001100070012000700130008001400090015000a0016000a001600 ,
                        0x0b0017000b0018000c0019000c001a000d001b000d001c000d001d000e001c00 ,
                        0x0e001b000d001a000d0019000d0018000c0017000c0016000c0015000b001400 ,
                        0x0b0013000b0012000a0011000a00100009000f0008000e0008000d0007000d00 ,
                        0x07000d0007000d0007000d0007000e0008000f00090010000a0011000b001200 ,
                        0x0b0012000c0013000c0014000d0015000d0016000e0017000e0018000f001900 ,
                        0x0f001a0010001a0010001b0010001a001000190010001800100018000f001700 ,
                        0x0f0016000f0015000f0014000e0013000e0012000d0011000d0010000d000f00 ,
                        0x0c000e000b000d000b000c000a000c000a000c000a000c000b000b000b000c00 ,
                        0x0b000d000c000e000d000f000e0010000e0011000f0012000f00130010001400 ,
                        0x10001500110016001100170012001700120018001200190013001a0013001900 ,
                        0x1300180013001700130016001200150012001400120013001100130011001200 ,
                        0x110011001000100010000f000f000e000f000d000f000b000e000a000d000a00 ,
                        0x0e000a000e000a000e000a000e000b000f000c000f000d0010000e0010000f00 ,
                        0x1100100012001100120012001300130013001400140015001400150014001600 ,
                        0x1500170015001800160019001600180016001700160016001500150015001400 ,
                        0x1500130015001200140011001400100014000f0013000e0013000d0012000c00 ,
                        0x12000b0012000a00110009001100090011000900110009001100090011000a00 ,
                        0x12000b0012000c0013000d0013000e0014000f00150010001500110016001200 ,
                        0x1600120016001300170014001700150018001600180017001800180019001700 ,
                        0x1900160018001500180014001800130018001200170011001700100017000f00 ,
                        0x17000e0016000d0016000c0015000b0015000a00150009001400080014000800 ,
                        0x1400080014000800140008001400090015000a0015000b0016000c0016000d00 ,
                        0x17000e0018000f0018000f001800100019001100190012001a0013001a001400 ,
                        0x1b0015001b0016001b0017001c0016001c0015001b0014001b0013001b001200 ,
                        0x1b0011001a0010001a000f001a000e001a000d0019000c0019000b0018000a00 ,
                        0x1800090018000800170007001700070017000600170006001700060017000700 ,
                        0x180008001800090019000a0019000b001a000c001b000d001b000e001b000f00 ,
                        0x1c0010001c0011001d0012001d0013001e0014001e0015001e0015001f001400 ,
                        0x1e0014001e0013001e0012001e0011001e0010001d000f001d000e001d000d00 ,
                        0x1d000c001c000b001c000a001b0009001b0007001b0006001a0005001a000500 ,
                        0x1a0005001a0005001a0005001a0006001b0007001b0008001c0009001c000a00 ,
                        0x1d000b001e000c001e000d001f000e001f000f001f0010002000110020001200 ,
                        0x2100130021001300210014002200130022001200210012002100110021001000 ,
                        0x21000f0021000e0020000d0020000c0020000a001f0009001f0008001f000700 ,
                        0x1e0006001e0005001d0004001d0004001d0003001d0003001d0003001d000500 ,
                        0x1e0006001e0007001f0008002000090020000a0021000b0021000c0022000d00 ,
                        0x22000e0023000f00230010002300110024001200240012002400130025001200 ,
                        0x250011002500100024000f0024000e0024000d0024000c0023000b0023000a00 ,
                        0x2300090023000800220007002200060021000400210003002000020020000200 ,
                        0x2000020020000200200002002000030021000400210005002200060023000800 ,
                        0x2300090024000a0025000b0025000c0025000d0026000e0026000f0027001000 ,
                        0x270011002700110028001200280011002800100028000f0028000e0027000d00 ,
                        0x27000c0027000b0027000a002700090026000800260007002500050025000400 ,
                        0x2400030024000200230000002300000023000000230000002300000023000200 ,
                        0x2400040025000600260007002700090027000a0028000c0029000d0029000f00 ,
                        0x290010002a0012002a0013002a0014002b0016002b0017002b0018002b001900 ,
                        0x2b001a002b001c002c001d002c001e002c001f002c0020002d0021002d002200 ,
                        0x2d0024002e0025002e0026002f0027003000280031002a0032002b0033002b00 ,
                        0x31002b002f002c002c002c002a002d0028002d0026002e0025002f0023002f00 ,
                        0x210030001f0031001d0032001c0034001a003500190036001700370016003600 ,
                        0x14003500130034001200330011003200100031000f002f000e002e000e002d00 ,
                        0x0d002c000d002b000d002a000c0028000c0027000c0026000b0025000b002400 ,
                        0x0b0022000b0021000a0020000a001f000a001d0009001c0009001b0008001900 ,
                        0x08001800070016000600150005001400040012000300110002000f0000000500 ,
                        0x000001020000000007000000fc020000000000000000040000002d0102000400 ,
                        0x0000f0010000040000002d0101008e0100002403c50037001c0037001c003700 ,
                        0x1b0037001b0037001b0036001b0036001b0036001b0036001a0036001a003600 ,
                        0x1a00350019003500180034001700340016003300150033001400320013003200 ,
                        0x1200310011003100100030000f0030000e002f000c002f000b002e000a002e00 ,
                        0x09002e0008002d0009002d000a002c000b002c000b002b000c002b000c002a00 ,
                        0x0d002a000d0029000e0028000e0028000f0027000f0026001000260010002500 ,
                        0x1000240011002300110023001100220011002200110022001100210011002100 ,
                        0x11002000110020001100200011001f0011001f0011001f0011001e0012001e00 ,
                        0x12001d0012001d0012001c0012001b0012001a00120019001200180012001800 ,
                        0x1200170013001600130015001300140013001300130012001300110014001000 ,
                        0x1400100014000f0014000e0015000d0015000c0015000b0016000b0016000a00 ,
                        0x1700090017000800170008001800070018000600190006001a0005001a000400 ,
                        0x1b0004001c0003001c0003001d0003001d0003001d0004001e0004001e000400 ,
                        0x1f00050020000500210006002200070024000700250008002600090028000900 ,
                        0x29000a002b000b002c000b002e000c0030000d0031000d0030000d0030000d00 ,
                        0x2f000d002f000d002f000e002e000e002e000e002d000f002d0010002c001000 ,
                        0x2c0011002b0012002a0013002a00140029001600290017002900180029001800 ,
                        0x29001800280019002800190028001a0028001a0028001b0028001b0028001c00 ,
                        0x28001c0028001c0028001d0028001d0028001e0028001e0027001e0027001e00 ,
                        0x27001e0027001e0028001e0029001e002a001f002b001f002c0020002d002000 ,
                        0x2e0021002f0021002f0022003000230031002400310025003200260032002700 ,
                        0x330028003300290033002b0033002c0033002d0032002e0032002f0031003000 ,
                        0x31003100300031002f0032002f0033002e0033002d0034002c0034002b003400 ,
                        0x2a00350029003500280035002700350026003400250034002400340023003300 ,
                        0x220033002100320020003300200033001f0034001f0034001e0035001e003600 ,
                        0x1d0036001d0037001c00050000000102ffffff0007000000fc020000ffffff00 ,
                        0x0000040000002d01000004000000f0010200040000002d0101006a0100002403 ,
                        0xb3003200200031001f0031001f0030001f0030001e002f001e002f001e002e00 ,
                        0x1d002e001d002d001d002d001d002c001d002c001d002b001c002b001c002a00 ,
                        0x1c0029001c0028001c0027001c0026001d0025001d0024001d0023001e002300 ,
                        0x1f0022001f002100200020002100200022001f0022001f0023001f0024001e00 ,
                        0x25001e0026001e0026001e0026001e0026001e0026001e0027001d0027001d00 ,
                        0x27001c0027001c0027001b0027001b0027001b0027001a0027001a0027001900 ,
                        0x2700190027001800270018002800170028001700280016002800150028001400 ,
                        0x2900130029001200290011002a0010002a0010002b000f002b000f002c000e00 ,
                        0x2c000e002d000d002d000d002e000d002e000d002e000c002d000c002c000b00 ,
                        0x2b000b002a000a0029000a002800090026000800250008002400070022000600 ,
                        0x21000600200005001f0005001e0004001d0004001d0005001c0005001b000600 ,
                        0x1b0006001a0007001a000800190008001900090018000a0018000b0017000b00 ,
                        0x17000c0016000d0016000e0016000e0016000f00150010001500110015001200 ,
                        0x1500130014001300140014001400150014001600140017001300180013001900 ,
                        0x13001a0013001b0013001b0013001c0013001d0013001e0013001e0013001e00 ,
                        0x12001f0012001f0012001f001200200012002000120021001200210012002100 ,
                        0x1200220012002200120023001200230012002300120023001200230012002400 ,
                        0x1200250011002600110026001100270010002800100029000f0029000f002a00 ,
                        0x0f002a000e002b000e002b000d002c000d002c000c002d000b002d000b002e00 ,
                        0x0c002e000d002f000e002f000f00300010003000110031001200310013003200 ,
                        0x140032001500330016003300170034001800340019003400190035001a003500 ,
                        0x1b0035001b0035001b0036001c0035001c0035001d0034001d0034001e003300 ,
                        0x1e0033001f0032001f0032002000050000000102ffffff0007000000fc020000 ,
                        0xffffff000000040000002d01020004000000f0010000040000002d0101001201 ,
                        0x0000240387002a002e002a002e0029002e0029002d0029002d0028002d002800 ,
                        0x2d0028002d0027002c0027002c0027002c0026002c0026002b0026002b002600 ,
                        0x2a0025002a0025002a002b002a002b0029002500290025002800250028002500 ,
                        0x280025002800250027002500270025002700250027002c0027002c0026002500 ,
                        0x2600250026002500250026002500260024002600240026002400270023002700 ,
                        0x23002700230028002200280022002900220029002200290022002a0022002a00 ,
                        0x22002b0022002b0022002c0022002d0022002d0022002e0023002e0023002e00 ,
                        0x24002f0023002f0023002f0022002e0022002e0022002e0022002e0022002d00 ,
                        0x21002d0021002d0021002c0021002c0021002c0021002b0021002b0021002b00 ,
                        0x21002a0021002a00210029002100290021002800210028002100270022002700 ,
                        0x2200260022002600230025002300250023002500240025002400240025002400 ,
                        0x2600240026002100260021002700240027002400270024002700240027002400 ,
                        0x2800240028002400280024002800240029002200290022002a0024002a002400 ,
                        0x2a0025002b0025002b0025002c0025002c0026002c0026002d0026002d002700 ,
                        0x2d0027002e0028002e0028002e0029002e0029002e002a002f002a002f002b00 ,
                        0x2f002c002e002c002e002d002e002d002e002e002d002e002d002f002d002e00 ,
                        0x2c002e002c002d002d002d002d002c002d002c002d002b002d002b002e002a00 ,
                        0x2e000500000001020000000007000000fc020000000000000000040000002d01 ,
                        0x000004000000f0010200040000002d01010086000000240341001d0015001e00 ,
                        0x15001e0014001f0014001f001400200014002000140020001300210013002100 ,
                        0x1300210012002200120022001200220011002200110022001000220010002200 ,
                        0x0f0022000f0022000e0022000e0022000d0021000d0021000c0021000c002000 ,
                        0x0c0020000c0020000b001f000b001f000b001e000b001e000b001d000b001d00 ,
                        0x0b001c000b001c000b001b000b001b000b001b000c001a000c001a000c001a00 ,
                        0x0c0019000d0019000d0019000e0019000e0018000f0018000f00180010001800 ,
                        0x100018001100190011001900120019001200190012001a0013001a0013001a00 ,
                        0x13001b0014001b0014001b0014001c0014001c0014001d0015001d0015000500 ,
                        0x00000102ffffff0007000000fc020000ffffff000000040000002d0102000400 ,
                        0x0000f0010000040000002d01010046000000240321001900100019000f001a00 ,
                        0x0e001a000d001b000d001b000c001c000c001c000c001d000c001e000c001f00 ,
                        0x0c0020000c0020000d0021000d0021000e0021000f0021001000210010002100 ,
                        0x11002100120020001200200013001f0013001e0014001d0014001c0014001c00 ,
                        0x13001b0013001b0012001a0012001a0011001900100019001000050000000102 ,
                        0x0000000007000000fc020000000000000000040000002d01000004000000f001 ,
                        0x0200040000002d010100860000002403410017001a0018001a0018001a001900 ,
                        0x1a0019001a001a0019001a0019001b0019001b0018001c0018001c0017001c00 ,
                        0x17001d0016001d0016001d0015001d0015001d0014001d0013001d0013001d00 ,
                        0x12001d0011001c0011001c0010001c0010001b000f001b000f001a000f001a00 ,
                        0x0e0019000e0019000e0018000e0018000e0017000e0016000e0016000e001500 ,
                        0x0e0015000e0014000e0014000f0013000f0013000f0012001000120010001200 ,
                        0x1100110011001100120011001300110013001100140011001500110015001100 ,
                        0x1600110016001200170012001700120018001300180013001900140019001400 ,
                        0x190015001a0015001a0016001a0016001a0017001a00050000000102ffffff00 ,
                        0x07000000fc020000ffffff000000040000002d01020004000000f00100000400 ,
                        0x00002d0101008600000024034100120014001200130012001300120012001200 ,
                        0x120012001100130011001300110013001000140010001400100015000f001500 ,
                        0x0f0015000f0016000f0016000f0017000f0018000f0018000f0019000f001900 ,
                        0x0f001a000f001a0010001a0010001b0010001b0011001b0011001c0011001c00 ,
                        0x12001c0012001c0013001c0013001c0014001c0014001c0015001c0015001c00 ,
                        0x16001c0016001b0017001b0017001b0018001a0018001a0018001a0018001900 ,
                        0x1900190019001800190018001900170019001600190016001900150019001500 ,
                        0x1900150018001400180014001800130018001300170013001700120016001200 ,
                        0x1600120015001200150012001400120014000500000001020000000007000000 ,
                        0xfc020000000000000000040000002d01000004000000f0010200040000002d01 ,
                        0x010046000000240321001c001d001c001d001d001d001e001c001e001c001f00 ,
                        0x1b001f001b001f001a001f0019001f0018001f0018001f0017001e0016001e00 ,
                        0x16001d0016001c0015001c0015001b0015001a00160019001600190016001800 ,
                        0x170018001800180018001800190018001a0018001b0018001b0019001c001900 ,
                        0x1c001a001d001b001d001c001d00050000000102ffffff0007000000fc020000 ,
                        0xffffff000000040000002d01020004000000f0010000040000002d0101004600 ,
                        0x000024032100190019001900190019001800190018001a0017001a0017001a00 ,
                        0x17001b0016001c0016001c0016001d0017001d0017001e0017001e0018001e00 ,
                        0x18001e0019001e0019001e001a001e001a001e001b001e001b001d001c001d00 ,
                        0x1c001c001c001c001c001b001c001a001c001a001c001a001b0019001b001900 ,
                        0x1a0019001a00190019000500000001020000000007000000fc02000000000000 ,
                        0x0000040000002d01000004000000f0010200040000002d010100ca0000002403 ,
                        0x63001e001e001d001f001d0020001c0022001c0023001b0024001b0025001b00 ,
                        0x26001b0027001b0029001b002a001c002c001c002d001d002e001d0030001e00 ,
                        0x31001f0032002000330021003400230034002400350025003500270036002800 ,
                        0x36002a0036002b0036002c0036002e0035002f00350030003400320034003300 ,
                        0x330034003200350031003600300036002e0037002d0037002c0038002a003800 ,
                        0x2900380027003800270038002600380025003800250038002400370023003700 ,
                        0x2300370022003600220036002300360024003700240037002500370026003700 ,
                        0x260037002700370027003700290037002a0036002c0036002d0035002e003500 ,
                        0x2f0034003000330031003200320031003300300033002f0034002e0034002c00 ,
                        0x35002b0035002a00350028003500270035002500340024003400230033002200 ,
                        0x330021003200200031001f0030001e002f001e002e001d002d001d002c001c00 ,
                        0x2a001c0029001c0027001c0026001c0025001c0024001d0023001d0022001e00 ,
                        0x21001e0020001f001f001e001e00030000000000
                    End
                    Picture ="euro.wmf"

                    TabIndex =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1320
                    Top =180
                    Width =6900
                    Height =480
                    Name ="lblComent"
                    Caption ="Bank details will be used in Conection with Customer and Suppliers. In order to "
                        "avoid duplicated records Bank \"Name\" must be uniqe."
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6060
                    Top =780
                    Width =1920
                    Height =300
                    FontWeight =700
                    Name ="lblPos"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =7200
                    Left =1320
                    Top =780
                    Width =4560
                    Height =300
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboFind"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblBankAccount.bnkaID, tblBankAccount.bnkaNo, tblBank.bnkName FROM tblBan"
                        "k INNER JOIN tblBankAccount ON tblBank.bnkID=tblBankAccount.bnkID WHERE (((tblBa"
                        "nkAccount.cstID)=1)) ORDER BY tblBankAccount.bnkaNo; "
                    ColumnWidths ="0;3600;3600"
                    AfterUpdate ="[Event Procedure]"
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =180
                    Height =255
                    ColumnWidth =1440
                    Name ="bnkaID"
                    ControlSource ="bnkaID"

                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1920
                    Top =1860
                    Width =3180
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="bnkaNo"
                    ControlSource ="bnkaNo"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =480
                            Top =1860
                            Width =1380
                            Height =255
                            ForeColor =128
                            Name ="bnkaNo_Label"
                            Caption ="Account No."
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2040
                    Top =3120
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="bnkaBLZ"
                    ControlSource ="bnkaBLZ"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =600
                            Top =3120
                            Width =1380
                            Height =255
                            Name ="bnkaBLZ_Label"
                            Caption ="BLZ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2040
                    Top =3480
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =9
                    Name ="bnkaABA"
                    ControlSource ="bnkaABA"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =600
                            Top =3480
                            Width =1380
                            Height =255
                            Name ="bnkaABA_Label"
                            Caption ="ABA"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2040
                    Top =3840
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =10
                    Name ="bnkaIBAN"
                    ControlSource ="bnkaIBAN"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =600
                            Top =3840
                            Width =1380
                            Height =255
                            Name ="bnkaIBAN_Label"
                            Caption ="BAN"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2040
                    Top =4200
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =11
                    Name ="bnkaSortingCode"
                    ControlSource ="bnkaSortingCode"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =540
                            Top =4200
                            Width =1440
                            Height =255
                            Name ="bnkaSortingCode_Label"
                            Caption ="SortingCode"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5400
                    Top =3120
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =12
                    Name ="bnkaSwiftCode"
                    ControlSource ="bnkaSwiftCode"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4440
                            Top =3120
                            Width =900
                            Height =255
                            Name ="bnkaSwiftCode_Label"
                            Caption ="SwiftCode"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5400
                    Top =3480
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =13
                    Name ="bnkaBeneficiary"
                    ControlSource ="bnkaBeneficiary"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4440
                            Top =3480
                            Width =900
                            Height =255
                            Name ="bnkaBeneficiary_Label"
                            Caption ="Beneficiary"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5400
                    Top =3840
                    Width =2280
                    Height =255
                    ColumnWidth =2310
                    TabIndex =14
                    Name ="bnkaIntermidiary"
                    ControlSource ="bnkaIntermidiary"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4440
                            Top =3840
                            Width =900
                            Height =255
                            Name ="bnkaIntermidiary_Label"
                            Caption ="Intermidiary"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =300
                    Top =300
                    Width =7680
                    Height =960
                    TabIndex =1
                    Name ="Frame24"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =480
                            Top =180
                            Width =540
                            Height =240
                            FontWeight =700
                            Name ="Label25"
                            Caption ="Bank"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =247
                    Left =300
                    Top =2760
                    Width =7680
                    Height =1980
                    TabIndex =7
                    Name ="Frame26"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =480
                            Top =2640
                            Width =1440
                            Height =240
                            FontWeight =700
                            Name ="Label27"
                            Caption ="Account Details"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2592
                    Left =6120
                    Top =1860
                    Width =1080
                    ColumnWidth =900
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"6\""
                    Name ="crnID"
                    ControlSource ="crnID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCurrency.crnID, tblCurrency.crnShortName, tblCurrency.crnName FROM tbl"
                        "Currency ORDER BY tblCurrency.crnShortName; "
                    ColumnWidths ="0;432;2160"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5280
                            Top =1860
                            Width =780
                            Height =255
                            ForeColor =128
                            Name ="crnID_Label"
                            Caption ="Currency"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =247
                    Left =300
                    Top =1440
                    Width =7680
                    Height =1140
                    TabIndex =4
                    Name ="Frame28"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =420
                            Top =1320
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label29"
                            Caption ="Account"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =7200
                    Left =1920
                    Top =840
                    Width =3840
                    ColumnWidth =900
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="bnkID"
                    ControlSource ="bnkID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblBank.bnkID, tblBank.bnkName, tblBank.bnkAddress FROM tblBank ORDER BY "
                        "tblBank.bnkName; "
                    ColumnWidths ="0;2880;4320"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =480
                            Top =840
                            Width =1380
                            Height =255
                            ForeColor =128
                            Name ="bnkID_Label"
                            Caption ="Bank"
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1920
                    Top =480
                    Width =5880
                    ColumnWidth =900
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"98\""
                    Name ="cstID"
                    ControlSource ="cstID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCustomer.cstID, tblCustomer.cstCompanyName FROM tblCustomer; "
                    ColumnWidths ="0;5760"
                    DefaultValue ="=[OpenArgs]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =480
                            Top =480
                            Width =1380
                            Height =255
                            ForeColor =128
                            Name ="cstID_Label"
                            Caption ="Customer/Supplier"
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
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =70
                    Left =360
                    Top =180
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
                    UnicodeAccessKey =70

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =76
                    Left =1620
                    Top =180
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
                    UnicodeAccessKey =76

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =78
                    Left =1200
                    Top =180
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
                    UnicodeAccessKey =78

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =80
                    Left =780
                    Top =180
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
                    UnicodeAccessKey =80

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =4800
                    Top =180
                    Width =1080
                    TabIndex =4
                    Name ="cmdAdd"
                    Caption ="&Add New"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add New Record"
                    UnicodeAccessKey =65

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =6000
                    Top =180
                    Width =1080
                    TabIndex =5
                    Name ="cmdDel"
                    Caption ="&Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =68

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =88
                    Left =7200
                    Top =180
                    Width =1080
                    TabIndex =6
                    Name ="cmdExit"
                    Caption ="E&xit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close Current Form"
                    UnicodeAccessKey =120

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


Private Sub cboFind_AfterUpdate()
On Error GoTo Err_End
' Find the record that matches the control.
If IsNoData(cboFind.Value) = False Then
    Me.Requery
    Me.RecordsetClone.FindFirst "[bnkaID] = " & Me![cboFind]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End If
Exit Sub
Err_End:
        Dim strMsg As String
        strMsg = FnIsErr(Err.Number)
        If strMsg <> "" Or IsNull(strMsg) Or IsEmpty(strMsg) Then
            MsgBox strMsg, vbExclamation, "Error!"
        Else
        End If
        cboFind.Value = ""
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
  bnkID.SetFocus
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
        bnkID.SetFocus
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
    DoCmd.Close
End Sub
Private Sub Form_Current()
' Purpose: Show current record position
If Me.NewRecord Then
    Me!lblPos.Caption = "Rec. New"
Else
    Me.RecordsetClone.Bookmark = Me.Bookmark
    Me!lblPos.Caption = "Rec. " & CStr(Me.RecordsetClone.AbsolutePosition + 1) _
                        & " of " & CStr(Me.RecordsetClone.RecordCount)
End If
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
Me.RecordSource = "SELECT tblBankAccount.* FROM tblBankAccount WHERE (((tblBankAccount.cstID)= " _
                  & Me.OpenArgs & "));"

cboFind.RowSource = "SELECT tblBankAccount.bnkaID, tblBankAccount.bnkaNo, tblBank.bnkName " _
                    & "FROM tblBank INNER JOIN tblBankAccount ON tblBank.bnkID=tblBankAccount.bnkID " _
                    & "WHERE (((tblBankAccount.cstID)=" & Me.OpenArgs & ")) ORDER BY tblBankAccount.bnkaNo;"
                  

End Sub
