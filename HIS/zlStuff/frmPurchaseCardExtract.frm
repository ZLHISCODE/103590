VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPurchaseCardExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提取数据"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "frmPurchaseCardExtract.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   14
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8040
      TabIndex        =   13
      Top             =   6120
      Width           =   1100
   End
   Begin TabDlg.SSTab sstGuide 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPurchaseCardExtract.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOption"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPurchaseCardExtract.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picView"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picView 
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3315
         ScaleWidth      =   3315
         TabIndex        =   15
         Top             =   120
         Width           =   3375
         Begin VSFlex8Ctl.VSFlexGrid vsfView 
            Height          =   2175
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2055
            _cx             =   3625
            _cy             =   3836
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.Frame fraOption 
         Caption         =   "提取设置"
         Height          =   2775
         Left            =   240
         TabIndex        =   1
         Top             =   150
         Width           =   5175
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   12
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   0
            Left            =   1800
            TabIndex        =   10
            Top             =   2160
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpData 
            Height          =   300
            Index           =   0
            Left            =   1800
            TabIndex        =   6
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   186580993
            CurrentDate     =   40532
         End
         Begin VB.OptionButton optExtract 
            Caption         =   "提取入库数据(&2)"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optExtract 
            Caption         =   "提取库存数据(&1)"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   480
            Value           =   -1  'True
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpData 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   8
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   186580993
            CurrentDate     =   40532
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   3240
            TabIndex        =   11
            Top             =   2210
            Width           =   180
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            Caption         =   "入库单号(&N)"
            Height          =   180
            Left            =   720
            TabIndex        =   9
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3240
            TabIndex        =   7
            Top             =   1730
            Width           =   180
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            Caption         =   "入库时间(&T)"
            Height          =   180
            Left            =   720
            TabIndex        =   5
            Top             =   1680
            Width           =   990
         End
         Begin VB.Label lblStock 
            AutoSize        =   -1  'True
            Caption         =   "库房： xxx"
            Height          =   180
            Left            =   720
            TabIndex        =   3
            Top             =   840
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmPurchaseCardExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngProviderID As Long  '供应商ID
Private mlngStockID As Long     '库房ID
Private mstrStock As String
Private mintUnit As Integer     '显示单位： 0-散装; 1-包装
Private mFMT As g_FmtString

Private Const mlngModule = 1712

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    If optExtract(0).Value Then
        Call ExtractStockData
    ElseIf optExtract(1).Value Then
        Call ExtractInStockData
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strReg As String
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(0, g_售价)
    End With
End Sub

Private Sub Form_Activate()
    Me.Visible = False
    Call cmd确定_Click
End Sub

Private Sub ExtractStockData()
    '提取库存数据
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select b.Id, '[' || b.编码 || ']' || b.名称 名称,e.名称 商品名, b.规格, b.产地,a.批准文号, " & IIf(mintUnit = 0, "b.计算单位", "c.包装单位") & " 单位," & vbNewLine & _
              "       Decode(b.是否变价, 1, a.实际金额 / a.实际数量, d.现价) * " & IIf(mintUnit = 0, "1", "c.换算系数") & " 售价," & vbNewLine & _
              "       Decode(c.在用分批, 1, a.上次采购价, (a.实际金额 - a.实际差价) / a.实际数量) * " & IIf(mintUnit = 0, "1", "c.换算系数") & " 成本价," & vbNewLine & _
              "       a.实际数量 / " & IIf(mintUnit = 0, "1", "c.换算系数") & " 数量," & vbNewLine & _
              "       c.最大效期, " & IIf(mintUnit = 0, "1", "c.换算系数") & " 换算系数, a.批次, b.是否变价, c.在用分批, c.指导差价率 / 100 指导差价率" & vbNewLine & _
              "From 药品库存 A, 收费项目目录 B, 材料特性 C, 收费价目 D, 收费项目别名 E" & vbNewLine & _
              "Where a.药品id = b.Id And a.药品id = c.材料id And a.药品id = d.收费细目id And a.性质 = 1 And a.库房id = [1] And a.实际数量 > 0 And" & vbNewLine & _
              "      a.上次供应商ID=[2] And b.类别 = '4' And b.撤档时间 >= To_Date('3000-1-1', 'yyyy-mm-dd') And d.终止日期 >= To_Date('3000-1-1', 'yyyy-mm-dd')" & vbNewLine & _
              GetPriceClassString("D") & " And b.Id = e.收费细目id(+) And e.性质(+) = 3" & vbNewLine & _
              " Order By a.药品ID, a.批次"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption, mlngStockID, mlngProviderID)
    FillData rsTmp
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExtractInStockData()
    '提取入库数据
End Sub

Private Sub FillData(ByVal rsVal As ADODB.Recordset)
    '填写数据到卡片中
    If rsVal.RecordCount = 0 Then
        MsgBox "未提取到库存数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim i As Long
    
    With frmPurchaseCard
        If .mshBill.Rows > 1 And Trim(.mshBill.TextMatrix(1, 0)) <> "" Then
            If MsgBox("退货卡片里有数据将全部清除后，再提取库存数据！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Me.MousePointer = vbHourglass
        .mshBill.Clear
        .mshBill.Rows = 2
        For i = 1 To rsVal.RecordCount
            .mshBill.TextMatrix(i, 1) = i
            .SetColValue i, rsVal!Id, rsVal!名称, IIf(IsNull(rsVal!规格), "", rsVal!规格), IIf(IsNull(rsVal!产地), "", rsVal!产地) _
                , IIf(IsNull(rsVal!单位), "", rsVal!单位) _
                , IIf(IsNull(rsVal!售价), 0, Format(rsVal!售价, IIf(mintUnit = 0, mFMT.FM_散装零售价, mFMT.FM_零售价))) _
                , IIf(IsNull(rsVal!成本价), 0, Format(rsVal!成本价, mFMT.FM_成本价)) _
                , IIf(IsNull(rsVal!产地), "", rsVal!产地), IIf(IsNull(rsVal!最大效期), 0, rsVal!最大效期), "" _
                , rsVal!换算系数, IIf(IsNull(rsVal!批次), 0, rsVal!批次), IIf(IsNull(rsVal!是否变价), 0, rsVal!是否变价) _
                , IIf(IsNull(rsVal!在用分批), 0, rsVal!在用分批), rsVal!指导差价率, IIf(IsNull(rsVal!批准文号), "", rsVal!批准文号), IIf(IsNull(rsVal!商品名), "", rsVal!商品名)
            .mshBill.TextMatrix(i, 21) = Format(rsVal!数量, mFMT.FM_数量)
            rsVal.MoveNext
            If Not rsVal.EOF Then .mshBill.Rows = .mshBill.Rows + 1
        Next
        Me.MousePointer = vbDefault
    End With
End Sub

Public Sub EntryPort(ByVal strStock As String, ByVal lngProviderID As Long)
    mlngStockID = Mid(strStock, 1, InStr(strStock, ";") - 1)
    mstrStock = Mid(strStock, InStr(strStock, ";") + 1)
    mlngProviderID = lngProviderID
    LblStock.Caption = "库房：" & mstrStock
End Sub

