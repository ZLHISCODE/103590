VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdviceSendDrugCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送条件"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmAdviceSendDrugCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabCond 
      Height          =   5535
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   564
      WordWrap        =   0   'False
      TabCaption(0)   =   "基本条件(&1)"
      TabPicture(0)   =   "frmAdviceSendDrugCond.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgLogo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTip(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDetail(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "给药途径(&2)"
      TabPicture(1)   =   "frmAdviceSendDrugCond.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetail(1)"
      Tab(1).Control(1)=   "lblTip(1)"
      Tab(1).Control(2)=   "imgLogo(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "药房置换(&3)"
      TabPicture(2)   =   "frmAdviceSendDrugCond.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDetail(2)"
      Tab(2).Control(1)=   "lblTip(2)"
      Tab(2).Control(2)=   "imgLogo(2)"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraDetail 
         Height          =   4425
         Index           =   2
         Left            =   -74835
         TabIndex        =   28
         Top             =   975
         Width           =   5400
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   4080
            Left            =   1170
            TabIndex        =   20
            Top             =   210
            Width           =   4095
            _cx             =   7223
            _cy             =   7197
            Appearance      =   1
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   6
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdviceSendDrugCond.frx":05DE
            ScrollTrack     =   -1  'True
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
            Editable        =   2
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
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院药房(&G)"
            Height          =   180
            Left            =   135
            TabIndex        =   19
            Top             =   300
            Width           =   990
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   4410
         Index           =   1
         Left            =   -74835
         TabIndex        =   26
         Top             =   975
         Width           =   5400
         Begin VB.CommandButton cmdAllWay 
            Caption         =   "全选"
            Height          =   330
            Left            =   180
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   3525
            Width           =   870
         End
         Begin VB.CommandButton cmdNoWay 
            Caption         =   "全清"
            Height          =   330
            Left            =   180
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   3900
            Width           =   870
         End
         Begin MSComctlLib.ListView lvwWay 
            Height          =   4050
            Left            =   1170
            TabIndex        =   16
            Top             =   210
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   7144
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "给药途径"
               Object.Width           =   6526
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "给药途径(&W)"
            Height          =   180
            Left            =   135
            TabIndex        =   15
            Top             =   300
            Width           =   990
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   4425
         Index           =   0
         Left            =   165
         TabIndex        =   24
         Top             =   975
         Width           =   5400
         Begin VB.CheckBox chkLimit 
            Caption         =   "给药途径执行以发送的结束时间为准计算"
            Height          =   360
            Left            =   3300
            TabIndex        =   5
            Top             =   510
            Width           =   1920
         End
         Begin VB.CheckBox chk加班加价 
            Caption         =   "执行加班加价(&V)"
            Height          =   195
            Left            =   3615
            TabIndex        =   6
            Top             =   585
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1260
            Width           =   4095
         End
         Begin VB.CommandButton cmdNoPati 
            Caption         =   "全清"
            Height          =   330
            Left            =   180
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   3570
            Width           =   870
         End
         Begin VB.CommandButton cmdAllPati 
            Caption         =   "全选"
            Height          =   330
            Left            =   180
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   3195
            Width           =   870
         End
         Begin VB.OptionButton opt期效 
            Caption         =   "临嘱(&T)"
            Height          =   180
            Index           =   1
            Left            =   2145
            TabIndex        =   2
            Top             =   255
            Width           =   930
         End
         Begin VB.OptionButton opt期效 
            Caption         =   "长嘱(&L)"
            Height          =   180
            Index           =   0
            Left            =   1170
            TabIndex        =   1
            Top             =   255
            Value           =   -1  'True
            Width           =   930
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1170
            TabIndex        =   4
            Top             =   540
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   64290819
            CurrentDate     =   37953
         End
         Begin MSComctlLib.ListView lvwPati 
            Height          =   2310
            Left            =   1170
            TabIndex        =   12
            Top             =   1620
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4075
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "姓名"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "住院号"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "床号"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "剩余款"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "住院医师"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "费别"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "护理等级"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "科室"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "入院日期"
               Object.Width           =   2857
            EndProperty
         End
         Begin VB.ComboBox cbo药房 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   900
            Width           =   4110
         End
         Begin MSComctlLib.Toolbar tbrAutoSel 
            Height          =   360
            Left            =   1170
            TabIndex        =   30
            Top             =   3990
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   635
            ButtonWidth     =   5318
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "按病区报警设置排开欠费病人   "
                  Object.ToolTipText     =   "Ctrl + Q"
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院病人(&P)"
            Height          =   180
            Left            =   135
            TabIndex        =   11
            Top             =   1695
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束时间(&E)"
            Height          =   180
            Left            =   135
            TabIndex        =   3
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院病区(&U)"
            Height          =   180
            Left            =   135
            TabIndex        =   9
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发药药房(&R)"
            Height          =   180
            Left            =   135
            TabIndex        =   7
            Top             =   960
            Width           =   990
         End
      End
      Begin VB.Label lblTip 
         Caption         =   "根据实际情况，将医嘱原来的执行药房指定为新的药房，医嘱发送时将会发送到新的药房执行。"
         Height          =   375
         Index           =   2
         Left            =   -73785
         TabIndex        =   29
         Top             =   585
         Width           =   4170
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Index           =   2
         Left            =   -74535
         Picture         =   "frmAdviceSendDrugCond.frx":0633
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTip 
         Caption         =   "可以通过选择不同的药品给药途径，来决定本次要发送哪些药品。"
         Height          =   375
         Index           =   1
         Left            =   -73785
         TabIndex        =   27
         Top             =   585
         Width           =   4170
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Index           =   1
         Left            =   -74535
         Picture         =   "frmAdviceSendDrugCond.frx":0EFD
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTip 
         Caption         =   "根据药品医嘱发送的需要，设置要发送的时间，医嘱类型；以及要发送医嘱的具体病人。"
         Height          =   375
         Index           =   0
         Left            =   1215
         TabIndex        =   25
         Top             =   585
         Width           =   4170
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Index           =   0
         Left            =   465
         Picture         =   "frmAdviceSendDrugCond.frx":17C7
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   555
      TabIndex        =   23
      Top             =   5715
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   22
      Top             =   5715
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3105
      TabIndex        =   21
      Top             =   5715
      Width           =   1100
   End
End
Attribute VB_Name = "frmAdviceSendDrugCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String 'IN
Public mlng病区ID As Long 'IN/OUT
Public mlng病人ID As Long 'IN
Public mblnOK As Boolean 'OUT:是否确认
Public mstrEnd As String 'OUT:结束时间
Public mint期效 As Integer 'OUT:0-长嘱,1-临嘱
Public mstr病人IDs As String 'OUT:病人ID串
Public mstr给药IDs As String 'OUT:给药途径ID串
Public mblnLimit As Boolean 'OUT:给药途径按发送结束时间限制计算
Public mlng药房ID As Long 'OUT:指定的药房
Public mrs药房 As ADODB.Recordset 'IN/OUT:药品替换集(可更新)

Private mrsWarn As ADODB.Recordset

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    
    strSQL = "Select 适用病人,报警方法,报警值 From 记帐报警线 Where 病区ID=[1]"
    Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    
    str病人IDs = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱发送病人", "")
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
        
    '在院病人:出院病人禁止下医嘱,发送医嘱
    strSQL = _
        "Select A.病人ID,A.姓名,A.住院号,B.出院病床 as 床号," & _
        " Nvl(E.预交余额,0)-Nvl(E.费用余额,0)+Decode(B.险类,Null,0,Nvl(F.金额,0)) as 剩余款," & _
        " A.担保额,Decode(X.编码,'1',1,Decode(B.险类,Null,0,1)) as 医保,B.险类," & _
        " B.住院医师,B.费别,D.名称 as 护理等级,C.名称 as 科室,B.入院日期" & _
        " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 D,病人余额 E,医疗付款方式 X," & _
        " (Select 病人ID,主页ID,Sum(金额) As 金额 From 保险模拟结算 Group By 病人ID,主页ID) F" & _
        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID=C.ID" & _
        " And A.病人ID=E.病人ID(+) And E.性质(+)=1 And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+)" & _
        " And B.出院日期 is NULL and Nvl(B.状态,0)<>3 And B.护理等级ID=D.ID(+) And B.医疗付款方式=X.名称(+)" & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) > 0, " And B.当前病区ID=[1]", "") & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) = 0, " Order by A.住院号 Desc", " Order by LPAD(B.出院病床,10,' ')")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
  
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!姓名)
        objItem.SubItems(1) = IIF(IsNull(rsTmp!住院号), "", rsTmp!住院号)
        objItem.SubItems(2) = IIF(IsNull(rsTmp!床号), "", rsTmp!床号)
        objItem.SubItems(3) = Format(Nvl(rsTmp!剩余款, 0), "0.00")
        objItem.SubItems(4) = IIF(IsNull(rsTmp!住院医师), "", rsTmp!住院医师)
        objItem.SubItems(5) = IIF(IsNull(rsTmp!费别), "", rsTmp!费别)
        objItem.SubItems(6) = IIF(IsNull(rsTmp!护理等级), "", rsTmp!护理等级)
        objItem.SubItems(7) = IIF(IsNull(rsTmp!科室), "", rsTmp!科室)
        objItem.SubItems(8) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
        
        '附加信息
        objItem.ListSubItems(1).Tag = Nvl(rsTmp!医保, 0)
        objItem.ListSubItems(2).Tag = Nvl(rsTmp!担保额, 0)
        
        '保险病人用红色显示
        If Not IsNull(rsTmp!险类) Then
            objItem.ForeColor = vbRed
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = vbRed
            Next
        End If
        
        '上次是否选择
        If cboUnit.ItemData(cboUnit.ListIndex) = lng病区ID And str病人IDs <> "" Then
            If InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 Then
                objItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objItem.EnsureVisible
                    objItem.Selected = True
                    k = 1
                End If
            End If
        ElseIf rsTmp!病人ID = mlng病人ID Then
            objItem.Checked = True '缺省只选择当前病人
            objItem.EnsureVisible
            objItem.Selected = True
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboUnit_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub chkLimit_Click()
    chkLimit.ForeColor = IIF(chkLimit.Value = 1, &HC0&, Me.ForeColor)
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdAllPati_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub cmdAllWay_Click()
    Call SelectLVW(lvwWay, True)
    lvwWay.SetFocus
End Sub

Private Sub cmdAllWay_GotFocus()
    tabCond.Tab = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdNoPati_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub cmdNoWay_Click()
    Call SelectLVW(lvwWay, False)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoWay_GotFocus()
    tabCond.Tab = 1
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "请选择一个病区。", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    '时间和期效
    mint期效 = IIF(opt期效(1).Value, 1, 0)
    If opt期效(0).Value Then
        mstrEnd = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Else
        mstrEnd = ""
    End If
    
    '给药途径计算方式
    mblnLimit = chkLimit.Value = 1
    
    '发药药房
    mlng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    
    '住院病人
    mstr病人IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mstr病人IDs = mstr病人IDs & "," & Mid(lvwPati.ListItems(i).Key, 2)
        End If
    Next
    mstr病人IDs = Mid(mstr病人IDs, 2)
    If mstr病人IDs = "" Then
        MsgBox "请至少选择一个需要发送医嘱病人。", vbInformation, gstrSysName
        tabCond.Tab = 0: lvwPati.SetFocus: Exit Sub
    End If
        
    '给药途径
    mstr给药IDs = ""
    For i = 1 To lvwWay.ListItems.Count
        If lvwWay.ListItems(i).Checked Then
            mstr给药IDs = mstr给药IDs & "," & Mid(lvwWay.ListItems(i).Key, 2)
        End If
    Next
    mstr给药IDs = Mid(mstr给药IDs, 2)
    If mstr给药IDs = "" Then
        MsgBox "请至少选择一种给药途径。", vbInformation, gstrSysName
        tabCond.Tab = 1: lvwWay.SetFocus: Exit Sub
    End If
    If UBound(Split(mstr给药IDs, ",")) + 1 = lvwWay.ListItems.Count Then
        mstr给药IDs = ""
    End If
    
    gbln加班加价 = chk加班加价.Value = 1
    
    mblnOK = True
    Unload Me
End Sub

Private Sub dtpEnd_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If tabCond.Tab = 0 Then
            Call cmdAllPati_Click
        Else
            Call cmdAllWay_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If tabCond.Tab = 0 Then
            Call cmdNoPati_Click
        Else
            Call cmdNoWay_Click
        End If
    ElseIf KeyCode = 13 Then
        If Not ActiveControl Is vsDept _
            And Not ActiveControl Is tabCond Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ActiveControl Is vsDept _
            And Not ActiveControl Is tabCond Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    
    '以发送结束时间为准
    chkLimit.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "限制结束时间", 0))
    
    '缺省医嘱期效
    lngTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱医嘱期效", 0))
    opt期效(lngTmp).Value = True
    
    '至少有一个才可能进来
    If InStr(mstrPrivs, "发送药疗临嘱") = 0 Then
        opt期效(0).Value = True
        opt期效(1).Enabled = False
    ElseIf InStr(mstrPrivs, "发送药疗长嘱") = 0 Then
        opt期效(1).Value = True
        opt期效(0).Enabled = False
    End If
   
    '缺省结束时间
    curDate = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱结束时点", "23:59:59")
    lngTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱时间间隔", 0))
    dtpEnd.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    
    '病区/病人
    Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
                        
    '发药药房
    Call Load药房
    
    '给药途径
    Call Load给药途径
    
    '药房置换
    Call Show药房
End Sub

Private Function Load药房() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    cbo药房.AddItem "所有药房"
    cbo药房.ListIndex = 0
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房')" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cbo药房.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo药房.ItemData(cbo药房.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    Load药房 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Show药房()
    Dim strTmp As String, i As Long, j As Long
    Dim str置换 As String, arr置换 As Variant
    
    mrs药房.Filter = 0
    If Not mrs药房.EOF Then
        vsDept.Rows = vsDept.FixedRows + mrs药房.RecordCount
        For i = 1 To mrs药房.RecordCount
            vsDept.Cell(flexcpData, i, 0) = CLng(mrs药房!ID)
            vsDept.TextMatrix(i, 0) = mrs药房!编码 & "-" & mrs药房!名称
            strTmp = strTmp & "|#" & mrs药房!ID & ";" & mrs药房!编码 & "-" & mrs药房!名称
            mrs药房.MoveNext
        Next
        
        str置换 = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱药房置换", "")
        arr置换 = Split(str置换, ",")
        For i = 1 To vsDept.Rows - 1
            mrs药房.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 0))
            For j = 0 To UBound(arr置换)
                If arr置换(j) Like mrs药房!ID & "-*" Then Exit For
            Next
            If j <= UBound(arr置换) Then
                mrs药房.Filter = "ID=" & Val(Split(arr置换(j), "-")(1))
                If Not mrs药房.EOF Then
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrs药房!ID)
                    mrs药房.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 0))
                    mrs药房!现ID = CLng(vsDept.Cell(flexcpData, i, 1))
                    mrs药房.Update
                Else
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrs药房!现ID)
                End If
            Else
                vsDept.Cell(flexcpData, i, 1) = CLng(mrs药房!现ID)
            End If
            
            mrs药房.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 1))
            vsDept.TextMatrix(i, 1) = mrs药房!编码 & "-" & mrs药房!名称
        Next
        If strTmp <> "" Then vsDept.ColComboList(1) = Mid(strTmp, 2)
    Else
        vsDept.Rows = vsDept.FixedRows + 1
        vsDept.Editable = flexEDNone
    End If
    vsDept.Row = vsDept.FixedRows: vsDept.Col = 1
    Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
End Sub

Private Function Load给药途径() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str给药IDs As String
    
    On Error GoTo errH
    
    str给药IDs = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱给药途径", "")
    
    strSQL = "Select ID,编码,名称 From 诊疗项目目录 Where 类别='E' And 操作类型='2' Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwWay.ListItems.Add(, "_" & rsTmp!ID, rsTmp!编码 & "-" & rsTmp!名称)
        
        If str给药IDs <> "" Then
            If InStr("," & str给药IDs & ",", "," & rsTmp!ID & ",") > 0 Then
                objItem.Checked = True
            End If
        Else
            objItem.Checked = True
        End If
        rsTmp.MoveNext
    Next
    Load给药途径 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        If Not gbln病区科室独立 Then
            strSQL = strSQL & IIF(strSQL <> "", " Union ", "") & _
                " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
                " From 床位状况记录 A,部门人员 B,部门表 C" & _
                " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
                " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        End If
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病区ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    
    '保存条件设置
    If mblnOK Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "限制结束时间", chkLimit.Value
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱结束时点", Format(dtpEnd.Value, "HH:mm:ss")
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱时间间隔", Int(CDate(Format(dtpEnd.Value, "yyyy-MM-dd")) - CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")))
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱医嘱期效", IIF(opt期效(1).Value, 1, 0)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱给药途径", mstr给药IDs
        
        '病人：选择病人仅为当前病人时,不保存
        If UBound(Split(mstr病人IDs, ",")) = 0 And Val(mstr病人IDs) = mlng病人ID Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱发送病人", ""
        Else
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱发送病人", cboUnit.ItemData(cboUnit.ListIndex) & ":" & mstr病人IDs
        End If
        
        '药房置换
        mrs药房.Filter = 0
        For i = 1 To mrs药房.RecordCount
            strTmp = strTmp & "," & mrs药房!ID & "-" & mrs药房!现ID
            mrs药房.MoveNext
        Next
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "药嘱药房置换", Mid(strTmp, 2)
    End If
    
    '释放私有及IN变量
    mstrPrivs = ""
    mlng病人ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub

Private Sub lvwPati_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub lvwWay_GotFocus()
    tabCond.Tab = 1
End Sub

Private Sub opt期效_Click(Index As Integer)
    dtpEnd.Enabled = opt期效(0).Value
    chkLimit.Visible = opt期效(0).Value
End Sub

Private Sub opt期效_GotFocus(Index As Integer)
    tabCond.Tab = 0
End Sub

Private Sub tabCond_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If tabCond.Tab = 0 Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf tabCond.Tab = 1 Then
            lvwWay.SetFocus
        ElseIf tabCond.Tab = 2 Then
            vsDept.SetFocus
        End If
    End If
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsDept.Cell(flexcpData, Row, Col) = CLng(vsDept.ComboData)
    mrs药房.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, Row, 0))
    mrs药房!现ID = CLng(vsDept.Cell(flexcpData, Row, Col))
    mrs药房.Update
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsDept.Editable <> flexEDNone And NewCol = 1 Then
        vsDept.FocusRect = flexFocusSolid
    Else
        vsDept.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsDept_GotFocus()
    tabCond.Tab = 2
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.Col = 1 Then
            If vsDept.Row + 1 <= vsDept.Rows - 1 Then
                vsDept.Row = vsDept.Row + 1
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                vsDept.Row = vsDept.FixedRows + 1
            End If
        Else
            vsDept.Col = 1
        End If
        Call vsDept.ShowCell(vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub vsDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.ComboIndex <> -1 Then
            Call vsDept_KeyPress(13)
        End If
    End If
End Sub

Private Sub vsDept_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                '只根据累计报警方法进行处理
                mrsWarn.Filter = "报警方法=1 And 适用病人=" & Val(.ListItems(i).ListSubItems(1).Tag) + 1
                If Not mrsWarn.EOF Then
                    If Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!报警值, 0) Then
                        .ListItems(i).Checked = False
                    End If
                End If
            End If
        Next
    End With
End Sub
