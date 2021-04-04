VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmChargeWholeSetItemEdit 
   Caption         =   "成套收费项目编辑"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmChargeWholeSetItemEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11745
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picUserDept 
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   495
      ScaleHeight     =   2460
      ScaleWidth      =   9855
      TabIndex        =   31
      Top             =   4620
      Width           =   9855
      Begin VB.TextBox txt科室 
         Height          =   300
         Left            =   495
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   45
         Width           =   3720
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "&L"
         Height          =   300
         Left            =   4200
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   5850
         TabIndex        =   23
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         Height          =   350
         Left            =   4710
         TabIndex        =   22
         Top             =   0
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvw科室 
         Height          =   4095
         Left            =   0
         TabIndex        =   24
         Top             =   420
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   7223
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "iltdept"
         SmallIcons      =   "iltdept"
         ForeColor       =   -2147483640
         BackColor       =   -2147483634
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   75
         TabIndex        =   19
         Top             =   105
         Width           =   360
      End
   End
   Begin VB.PictureBox picWholeSet 
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   675
      ScaleHeight     =   2040
      ScaleWidth      =   11145
      TabIndex        =   30
      Top             =   3270
      Width           =   11145
      Begin VSFlex8Ctl.VSFlexGrid vsWholeSet 
         Height          =   4680
         Left            =   195
         TabIndex        =   18
         Top             =   195
         Width           =   11355
         _cx             =   20029
         _cy             =   8255
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeWholeSetItemEdit.frx":0442
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
         ExplorerBar     =   2
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
   Begin VB.PictureBox picCmd 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11745
      TabIndex        =   28
      Top             =   7770
      Width           =   11745
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   26
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9330
         TabIndex        =   25
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fra成套 
      Caption         =   "成套项目信息"
      Height          =   1950
      Left            =   120
      TabIndex        =   27
      Top             =   255
      Width           =   11460
      Begin VB.TextBox txtMemo 
         Height          =   300
         Left            =   900
         MaxLength       =   100
         TabIndex        =   17
         Tag             =   "备注"
         Top             =   1485
         Width           =   3795
      End
      Begin VB.ComboBox cbo人员 
         Height          =   300
         Left            =   5895
         TabIndex        =   15
         Text            =   "cbo人员"
         Top             =   1110
         Width           =   4005
      End
      Begin VB.OptionButton opt范围 
         Caption         =   "本院"
         Height          =   315
         Index           =   2
         Left            =   3255
         TabIndex        =   13
         Top             =   1125
         Width           =   1530
      End
      Begin VB.OptionButton opt范围 
         Caption         =   "指定科室"
         Height          =   315
         Index           =   1
         Left            =   2025
         TabIndex        =   12
         Top             =   1110
         Width           =   1530
      End
      Begin VB.OptionButton opt范围 
         Caption         =   "指定人员"
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   11
         Top             =   1110
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.TextBox txtWB 
         Height          =   300
         Left            =   7860
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "五笔"
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox txtParent 
         Height          =   300
         Left            =   5895
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "(无)"
         ToolTipText     =   "按Del清除上级，设置初级分类"
         Top             =   285
         Width           =   3720
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&P"
         Height          =   300
         Left            =   9600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   285
         Width           =   285
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   915
         MaxLength       =   100
         TabIndex        =   6
         Tag             =   "名称"
         Top             =   750
         Width           =   3795
      End
      Begin VB.TextBox txtSymbol 
         Height          =   300
         Left            =   5895
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "拼音"
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   930
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "编码"
         Text            =   "0000"
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "备注(&M)"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label lbl人员 
         AutoSize        =   -1  'True
         Caption         =   "指定人员"
         Height          =   180
         Left            =   5070
         TabIndex        =   14
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "使用范围"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblParent 
         AutoSize        =   -1  'True
         Caption         =   "分类(&U)"
         Height          =   180
         Left            =   5175
         TabIndex        =   2
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   795
         Width           =   630
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         Caption         =   "编码(&D)"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lblSymbol 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)                 (拼音)                (五笔)"
         Height          =   180
         Left            =   5160
         TabIndex        =   7
         Top             =   780
         Width           =   4680
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   4110
      Left            =   105
      TabIndex        =   32
      Top             =   2355
      Width           =   11595
      _Version        =   589884
      _ExtentX        =   20452
      _ExtentY        =   7250
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList iltdept 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeWholeSetItemEdit.frx":074A
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChargeWholeSetItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditTypeWhileSetItem
    EdI_增加 = 1
    EdI_修改 = 2
    EdI_查看 = 3
End Enum
Private mEditType As EditTypeWhileSetItem
Private mstrPrivs As String, mlngModule As Long
Private mstrWholeItems As String
Private mstrID As String, mlng分类ID As Long
Private mintSucces As Integer
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSort As Boolean
Private mbln修改 As Boolean
Private Enum mItemPage
    pg_成套组合 = 1
    pg_使用科室 = 2
End Enum
Private mrsDept As ADODB.Recordset
Private Sub zlInitClassPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分类页面
    '编制:刘兴洪
    '日期:2010-08-24 10:15:11
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Err = 0: On Error GoTo ErrHand:
 
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_成套组合, "成套项目组成(&1)", picWholeSet.hwnd, 0)
    ObjItem.Tag = mItemPage.pg_成套组合
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_使用科室, "执行科室(&2)", picUserDept.hwnd, 0)
    ObjItem.Tag = mItemPage.pg_使用科室
    tbPage.Item(0).Selected = True
     With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function ShowCard(ByVal frmMain As Form, ByVal EditType As EditTypeWhileSetItem, ByVal strPrivs As String, ByVal lngModule As Long, _
    Optional lng分类id As Long = 0, Optional strID As String = "", Optional strWholeItems As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,显示或编辑相关的信息数据
    '入参:frmMain-调用主窗口
    '       EditType-编辑类型
    '       strWholeItems-从记帐单中传入的单据数据,格式为:
    '                              序号,父号,收费细目ID,数量,单价,执行科室|序号,父号,收费细目ID,数量,单价,执行科室|…
    '出参:strID-返回当前编辑的ID
    '返回:
    '编制:刘兴洪
    '日期:2010-08-26 17:00:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule: mlng分类ID = lng分类id: mstrID = strID
    mstrWholeItems = strWholeItems: mintSucces = 0
    Me.Show 1, frmMain
    ShowCard = mintSucces > 0
    strID = mstrID
End Function
Private Sub InitDefaultLen()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化默认数据库长度
    '编制:刘兴洪
    '日期:2010-08-26 17:08:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "" & _
    " Select A.ID,A.分类ID,A.编码,A.名称,A.拼音,A.五笔,A.备注" & _
    " From 成套收费项目 A" & _
    " Where id=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
    Me.txtCode.MaxLength = rsTemp.Fields("编码").DefinedSize
    Me.txtName.MaxLength = rsTemp.Fields("名称").DefinedSize
    Me.txtSymbol.MaxLength = rsTemp.Fields("拼音").DefinedSize
    Me.txtWB.MaxLength = rsTemp.Fields("五笔").DefinedSize
    Me.txtMemo.MaxLength = rsTemp.Fields("备注").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AnalyzeWholeSetItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:分解传入的成套项目数据
    '编制:刘兴洪
    '日期:2010-08-26 17:17:50
    '说明:mstrWholeItems的格式为:序号,父号,收费细目ID,付数,数量,单价,执行科室|序号,父号,收费细目ID,数量,单价,执行科室|…
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long, j As Long, m As Long, lngId As Long
    Dim strValue(0 To 10) As String, strSubItem As String, str收费细目ID As String, str执行科室ID As String
    Dim strDeptValue(0 To 10) As String, strDeptSub As String, lng父号 As Long
    Dim rsItems As ADODB.Recordset, rsDept As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng主项ID As Long
    Dim cllTemp As Collection
    
    If mstrWholeItems = "" Then Exit Sub
    
    On Error GoTo ErrHandle
    
    '先分解出来,再查找
    varData = Split(mstrWholeItems, "|")
    For i = 0 To UBound(varData)
        '序号,父号,收费细目ID,付数,数量,单价,执行科室
        varTemp = Split(varData(i) & ",,,,,", ",")
        If Len(str收费细目ID) > 1990 And j <= 10 Then
            strValue(j) = Mid(str收费细目ID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as 收费细目ID From Table(f_Num2List([" & j + 1 & "])) B "
            str收费细目ID = "": j = j + 1
        End If
        str收费细目ID = str收费细目ID & "," & Val(varTemp(2))
        If Len(str执行科室ID) > 1990 And m <= 10 Then
            strDeptValue(m) = Mid(str执行科室ID, 2)
            strDeptSub = strDeptSub & " Union ALL " & _
            " Select Column_Value as 执行部门ID From Table(f_Num2List([" & m + 1 & "])) B "
            m = m + 1
            str执行科室ID = ""
        Else
            str执行科室ID = str执行科室ID & "," & Val(varTemp(6))
        End If
    Next
    If str收费细目ID <> "" Then
        If j > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From 收费项目目录 Where id in (" & Mid(str收费细目ID, 2) & ")"
        Else
            strValue(j) = Mid(str收费细目ID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as 收费细目ID From Table(f_Num2List([" & j + 1 & "])) B "
        End If
    End If
    If str执行科室ID <> "" Then
        If m > 10 Then
             strDeptSub = strDeptSub & " UNION ALL Select ID From 部门表 Where id in (" & Mid(str执行科室ID, 2) & ")"
        Else
            strDeptValue(m) = Mid(str执行科室ID, 2)
            strDeptSub = strDeptSub & " Union ALL " & _
            " Select Column_Value as 执行部门ID From Table(f_Num2List([" & m + 1 & "])) B "
        End If
    End If
    
    gstrSQL = "" & _
       "   Select A.主项id, A.从项id, A.固有从属, A.从项数次 " & _
       "   From 收费从属项目 A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.主项id = D.收费细目id"
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    
    gstrSQL = "" & _
    "   Select A.类别, A.ID,A.编码,A.名称,A.计算单位,A.规格,B.中药形态,c.编码 as 诊疗编码," & _
    "             C.名称 as 诊疗名称,C.计算单位 as 剂量单位,B.药名Id,B.剂量系数,A.执行科室,A.是否变价, B1.跟踪在用" & _
    "   From 收费项目目录 A,药品规格 B,材料特性 B1,诊疗项目目录 C,(" & Mid(strSubItem, 11) & ") D" & _
    "   Where A.ID=b.药品ID(+) and A.ID=b1.材料ID(+) and B.药名Id=C.ID(+) and A.id=d.收费细目ID "
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    If strDeptSub <> "" Then
        gstrSQL = "" & _
        "   Select A.ID,A.编码,A.名称 " & _
        "   From 部门表 A,(" & Mid(strDeptSub, 11) & ") D" & _
        "   Where A.id =D.执行部门ID"
    Else
        gstrSQL = "" & _
        "   Select A.ID,A.编码,A.名称 " & _
        "   From 部门表 A " & _
        "   Where A.id =0"
    End If
    Set rsDept = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strDeptValue(0), strDeptValue(1), strDeptValue(2), strDeptValue(3), strDeptValue(4), strDeptValue(5), strDeptValue(6), strDeptValue(7), strDeptValue(8), strDeptValue(9), strDeptValue(10))
    With vsWholeSet
        .Clear 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("标志"): .SubtotalPosition = flexSTAbove
        .Rows = IIF(UBound(varData) = 0, 1, UBound(varData) + 1) + 1
        str收费细目ID = "": str执行科室ID = "": j = 0: m = 0
        Set cllTemp = New Collection
        
        For i = 0 To UBound(varData)
            '序号,父号,收费细目ID,付数,数量,单价,执行科室|
            varTemp = Split(varData(i) & ",,,,,", ",")
            .TextMatrix(i + 1, .ColIndex("序号")) = i + 1
            .Cell(flexcpData, i + 1, .ColIndex("序号")) = i + 1 ' Val(varTemp(0))
            cllTemp.Add i + 1, "_" & Val(varTemp(0))
            
            If Val(varTemp(1)) = 0 Then
                .TextMatrix(i + 1, .ColIndex("从属父号")) = ""
            Else
                .TextMatrix(i + 1, .ColIndex("从属父号")) = cllTemp("_" & Val(varTemp(1)))
            End If
            .Cell(flexcpData, i + 1, .ColIndex("收费项目")) = Val(varTemp(2))
            
            .TextMatrix(i + 1, .ColIndex("缺省付数")) = IIF(Val(varTemp(3)) = 0, 1, Val(varTemp(3)))
            
            .TextMatrix(i + 1, .ColIndex("缺省数量")) = FormatEx(Val(varTemp(4)), 5)
            .Cell(flexcpData, i + 1, .ColIndex("缺省数量")) = Val(varTemp(4))
            .TextMatrix(i + 1, .ColIndex("缺省价格")) = FormatEx(Val(varTemp(5)), 8)
            .Cell(flexcpData, i + 1, .ColIndex("缺省价格")) = Val(varTemp(5))
            .Cell(flexcpData, i + 1, .ColIndex("缺省执行科室")) = Val(varTemp(6))
            
            lngId = Val(.Cell(flexcpData, i + 1, .ColIndex("收费项目")))
            If Val(.TextMatrix(i + 1, .ColIndex("从属父号"))) = 0 Then
                lng主项ID = lngId
            End If
            
            rsItems.Find "ID=" & lngId, , adSearchForward, 1
            If rsItems.EOF = False Then
                .TextMatrix(i + 1, .ColIndex("收费项目")) = NVL(rsItems!编码) & "-" & NVL(rsItems!名称)
                .TextMatrix(i + 1, .ColIndex("规格")) = NVL(rsItems!规格)
                .TextMatrix(i + 1, .ColIndex("药名ID")) = NVL(rsItems!药名ID)
                .TextMatrix(i + 1, .ColIndex("跟踪在用")) = Val(NVL(rsItems!跟踪在用))
                If NVL(rsItems!类别) = "7" Then
                    '草药,显示诊疗名称
                    .TextMatrix(i + 1, .ColIndex("药名")) = NVL(rsItems!诊疗编码) & "-" & NVL(rsItems!诊疗名称)
                    .TextMatrix(i + 1, .ColIndex("单位")) = NVL(rsItems!剂量单位)
                    .TextMatrix(i + 1, .ColIndex("缺省数量")) = FormatEx(Val(.TextMatrix(i + 1, .ColIndex("缺省数量"))) * Val(NVL(rsItems!剂量系数)), 5)
                    .TextMatrix(i + 1, .ColIndex("缺省价格")) = FormatEx(Val(.TextMatrix(i + 1, .ColIndex("缺省价格"))) / Val(NVL(rsItems!剂量系数)), 8)
                    .TextMatrix(i + 1, .ColIndex("中药形态")) = Val(NVL(rsItems!中药形态))
                    .TextMatrix(i + 1, .ColIndex("剂量系数")) = Val(NVL(rsItems!剂量系数))
                    
'                    If Val(Nvl(rsItems!中药形态)) = 0 Then   '散装形态的,则显示具体的规格
'                        .TextMatrix(i + 1, .ColIndex("规格")) = Nvl(rsItems!规格)
'                    End If
                Else
                    .TextMatrix(i + 1, .ColIndex("单位")) = NVL(rsItems!计算单位)
                End If
                .TextMatrix(i + 1, .ColIndex("类别")) = NVL(rsItems!类别)
                .TextMatrix(i + 1, .ColIndex("是否变价")) = NVL(rsItems!是否变价)
                .TextMatrix(i + 1, .ColIndex("执行科室")) = NVL(rsItems!执行科室)
            End If
            rsDept.Find "ID=" & Val(.Cell(flexcpData, i + 1, .ColIndex("缺省执行科室"))), , adSearchForward, 1
            If Not rsDept.EOF Then
                .TextMatrix(i + 1, .ColIndex("缺省执行科室")) = NVL(rsDept!编码) & "-" & NVL(rsDept!名称)
            End If
            
            If Not rsOthers Is Nothing And Val(.TextMatrix(i + 1, .ColIndex("从属父号"))) <> 0 Then
                '  "   Select A.主项id, A.从项id, A.固有从属, A.从项数次 "
                rsOthers.Filter = "主项ID=" & lng主项ID & " And 从项ID= " & Val(.Cell(flexcpData, i + 1, .ColIndex("收费项目")))
                If Not rsOthers.EOF Then
                    .TextMatrix(i + 1, .ColIndex("从属数次")) = Val(NVL(rsOthers!从项数次))
                    .Cell(flexcpData, i + 1, .ColIndex("从属父号")) = Val(NVL(rsOthers!固有从属))
                End If
            End If
            
            If Val(.TextMatrix(i + 1, .ColIndex("从属父号"))) = 0 Then
                    lng父号 = i + 1 'Val(.Cell(flexcpData, i + 1, .ColIndex("序号")))
                    .IsSubtotal(i + 1) = True: .RowOutlineLevel(i + 1) = 1
            ElseIf lng父号 = Val(.TextMatrix(i + 1, .ColIndex("从属父号"))) Then
                If i + 1 > 2 Then
                   If Val(.TextMatrix(i, .ColIndex("从属父号"))) <> 0 Then
                        .IsSubtotal(i) = False
                        .RowOutlineLevel(i) = 2
                    End If
                End If
                .IsSubtotal(i + 1) = True
                .RowOutlineLevel(i + 1) = 2
            End If
        Next
    End With
 
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ClearCardData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检除卡片数据
    '编制:刘兴洪
    '日期:2010-08-27 10:01:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtCode.Text = "": txtName.Text = ""
    txtMemo.Text = "": txtSymbol = ""
    txtWB.Text = ""
    cbo人员.Text = "": txt科室.Text = ""
    cbo人员.ListIndex = -1
    lvw科室.ListItems.Clear
    vsWholeSet.Clear 1: vsWholeSet.Rows = 2
    vsWholeSet.TextMatrix(1, vsWholeSet.ColIndex("序号")) = 1
    vsWholeSet.ColWidth(vsWholeSet.ColIndex("标志")) = 240
End Sub

Private Sub EditStatusSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置编辑状态
    '编制:刘兴洪
    '日期:2010-08-27 10:24:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control
    If mEditType = EdI_增加 Then
        With vsWholeSet
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = .ColIndex("标志"): .SubtotalPosition = flexSTAbove
            .Editable = flexEDKbdMouse
        End With
            opt范围(1).Enabled = InStr(1, mstrPrivs, ";本科成套方案;") > 0
            opt范围(2).Enabled = InStr(1, mstrPrivs, ";全院成套方案;") > 0
'            If opt范围(1).Enabled And opt范围(1).value = True Then
'               opt范围(0).value = True
'            End If
'            If opt范围(2).Enabled And opt范围(2).value = True Then
'               opt范围(0).value = True
'            End If
            txtCode.Enabled = True
    ElseIf mEditType = EdI_修改 And mbln修改 Then
        
        With vsWholeSet
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = .ColIndex("标志"): .SubtotalPosition = flexSTAbove
            .Editable = flexEDKbdMouse
        End With
            opt范围(1).Enabled = InStr(1, mstrPrivs, ";本科成套方案;") > 0
            opt范围(2).Enabled = InStr(1, mstrPrivs, ";全院成套方案;") > 0
'            If opt范围(1).Enabled And opt范围(1).value = True Then
'               opt范围(0).value = True
'            End If
'            If opt范围(2).Enabled And opt范围(2).value = True Then
'               opt范围(0).value = True
'            End If
            txtCode.Enabled = True
    Else
        With vsWholeSet
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = .ColIndex("序号"): .SubtotalPosition = flexSTAbove
            .Editable = flexEDNone
        End With
        
        For Each objCtl In Me.Controls
            Select Case UCase(TypeName(objCtl))
            Case "TEXTBOX"
                objCtl.Enabled = False
                objCtl.BackColor = Me.BackColor
            Case UCase("OptionButton")
                objCtl.Enabled = False
            Case UCase("ComBox")
                objCtl.Enabled = False
                objCtl.BackColor = Me.BackColor
            Case UCase("CommandButton")
                If Not (objCtl Is cmdOK Or objCtl Is cmdCancel Or objCtl Is cmdHelp) Then
                        objCtl.Enabled = False
                End If
            Case UCase("vsFlexGrid")
                objCtl.Editable = flexEDNone
            Case Else
            End Select
        Next
        
        Me.cbo人员.Enabled = False
        cbo人员.BackColor = Me.BackColor
    End If
End Sub

Private Sub zlDefaultCode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省编码
    '编制:刘兴洪
    '日期:2010-08-27 12:00:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngLen As String, strUpCode As String
    On Error GoTo ErrHandle
    If Val(txtParent.Tag) = 0 Then
NotNO:
        gstrSQL = "Select Max(编码) as 编码 From 成套项目分类  Where 上级ID is null  "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        txtParent.Text = "无": txtParent.Tag = ""
        If NVL(rsTemp!编码) = "" Then
            txtCode.Text = "001"
        Else
            strTemp = Val(rsTemp!编码) + 1
            If Len(strTemp) > Len(rsTemp!编码) Then
                txtCode.Text = strTemp
            Else
                 txtCode.Text = String(Len(rsTemp!编码) - Len(strTemp), "0") & strTemp
            End If
        End If
        Exit Sub
    End If
    
    gstrSQL = "Select ID,编码,名称 From 成套项目分类  Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtParent.Tag))
    If rsTemp.EOF Then
        GoTo NotNO:
    End If
    
    strUpCode = NVL(rsTemp!编码)
    txtParent.Text = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
    txtParent.Tag = NVL(rsTemp!ID)
    gstrSQL = "select max(编码) as 编码" & _
            " From 成套收费项目" & _
            " Where   编码 like [1] And 编码<> [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strUpCode & "%", strUpCode)
    If rsTemp.EOF Then
         txtCode.Text = strUpCode & "01"
    Else
        strTemp = NVL(rsTemp!编码)
        If strTemp = "" Then
            txtCode.Text = strUpCode & "01"
        Else
            txtCode.Text = Val(strTemp) + 1
            lngLen = Len(strTemp) - Len(txtCode)
            If lngLen > 0 Then
                txtCode.Text = String(lngLen, "0") & txtCode.Text
            End If
         End If
   
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function ReadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取项目信息
    '返回:读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-26 13:35:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lngRow As Long, lng父号 As Long, i As Long, j As Long
    Dim lng主项ID As Long
    
    Call EditStatusSet  '设置编辑状态
    If mEditType = EdI_增加 Then
        Call ClearCardData
        txtParent.Tag = mlng分类ID
        Call zlDefaultCode
        If mstrWholeItems <> "" Then Call AnalyzeWholeSetItem
        Call opt范围_Click(0)
        ReadCard = True
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    Call ClearCardData
    
    On Error GoTo ErrHandle
    If mEditType <> EdI_查看 Then
        gstrSQL = "" & _
         "   Select A.主项id, A.从项id, A.固有从属, A.从项数次 " & _
         "   From 收费从属项目 A, 成套收费项目组合 B " & _
         "   Where 主项id = B.收费细目id And Nvl(B.从属父号, 0) = 0 And B.成套id = [1]"
        Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
    Else
        Set rsOthers = Nothing
    End If
    
    Dim strWherePriceGrade As String
    If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
        strWherePriceGrade = " And j.价格等级 Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And j.价格等级 = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And j.价格等级 = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And j.价格等级 = [4])" & vbNewLine & _
            "      Or (j.价格等级 Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From 收费价目" & vbNewLine & _
            "                          Where j.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And 价格等级 = [4])))))"
    End If
    gstrSQL = "" & _
    "   Select  /*+Rule */ A.成套ID,A.收费细目ID,A.序号,A.从属父号,A.付数,A.数量,A.单价,A.执行科室ID, " & _
    "              B.类别,B.编码,B.名称,B.计算单位,B.规格,C.中药形态,D.编码 as 诊疗编码, " & _
    "              D.名称 as 诊疗名称,D.计算单位 as 剂量单位,C.剂量系数, " & _
    "              E.编码 As 执行科室编码,E.名称 As 执行科室名称, " & _
    "              M.编码 As 成套编码,M.名称 As 成套名称,M.拼音,M.五笔,M.备注,M.范围, " & _
    "              M.分类ID,M.人员ID,G.姓名,J.编码 As 分类编码,J.名称 As 分类名称 ,B.是否变价,B.执行科室,C.药名ID,B1.跟踪在用," & _
    "              Decode(B.是否变价,1,'时价',LTrim(To_Char(J1.现价,'999999999.9999999'))) as 现价" & _
    "   From 成套收费项目 M,成套项目分类 J,成套收费项目组合 A,收费项目目录 B,材料特性 B1,药品规格 C,诊疗项目目录 D, " & _
    "             部门表 E,人员表 G, " & _
    "             (Select j.收费细目id, Sum(j.现价) as 现价" & vbNewLine & _
    "              From 收费价目 J,收费项目目录 K" & vbNewLine & _
    "              Where j.收费细目ID = k.ID And Sysdate Between J.执行日期 And Nvl(J.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.收费细目id ) J1 " & _
    "   Where   M.分类ID=J.Id And  M.人员ID=G.Id(+) And M.Id=A.成套ID(+)  " & _
    "               And A.收费细目id=J1.收费细目ID(+)" & _
    "               And A.收费细目id=b.Id(+)  and a.收费细目ID=b1.材料ID(+) And a.收费细目ID=C.药品ID(+) And C.药名ID=D.Id(+) " & _
    "               And A.执行科室ID=E.Id(+)  " & _
    "               And M.ID=[1] " & _
    "   Order by A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID), gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    
    If rsTemp.EOF Then
        MsgBox "该成套收费项目可能已经被他人删除,不能进行修改或查看!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    mlng分类ID = Val(NVL(rsTemp!分类id))
    txtParent.Text = IIF(mlng分类ID = 0, "无", NVL(rsTemp!分类编码) & "-" & NVL(rsTemp!分类名称))
    txtParent.Tag = mlng分类ID
    txtCode.Text = NVL(rsTemp!成套编码)
    txtCode.Tag = NVL(rsTemp!成套编码)
    txtName.Text = NVL(rsTemp!成套名称)
    txtName.Tag = NVL(rsTemp!成套名称)
    txtSymbol.Text = NVL(rsTemp!拼音)
    txtWB.Text = NVL(rsTemp!五笔)
    txtMemo.Text = NVL(rsTemp!备注)
    
     '0-全院;1-科室;2-操作员
    If Val(NVL(rsTemp!范围)) = 0 Then '全院
        opt范围(2).value = True
    ElseIf Val(NVL(rsTemp!范围)) = 1 Then
        opt范围(1).value = True
    Else
        opt范围(0).value = True
        cbo人员.AddItem NVL(rsTemp!姓名)
        cbo人员.ItemData(cbo人员.NewIndex) = Val(NVL(rsTemp!人员ID))
        cbo人员.ListIndex = cbo人员.NewIndex
        cbo人员.Tag = cbo人员.ListIndex
    End If
    
    '加载组成
    With vsWholeSet
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("标志"): .SubtotalPosition = flexSTAbove
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("序号")) = IIF(Val(NVL(rsTemp!序号)) = 0, lngRow, Val(NVL(rsTemp!序号)))
            .Cell(flexcpData, lngRow, .ColIndex("序号")) = Val(NVL(rsTemp!序号))
            .TextMatrix(lngRow, .ColIndex("从属父号")) = Val(NVL(rsTemp!从属父号))
            .Cell(flexcpData, lngRow, .ColIndex("收费项目")) = Val(NVL(rsTemp!收费细目id))
            .TextMatrix(lngRow, .ColIndex("缺省付数")) = IIF(Val(NVL(rsTemp!付数)) = 0, 1, Val(NVL(rsTemp!付数)))
            .TextMatrix(lngRow, .ColIndex("缺省数量")) = FormatEx(Val(NVL(rsTemp!数量)), 5)
            .Cell(flexcpData, lngRow, .ColIndex("缺省数量")) = Val(NVL(rsTemp!数量))
            .TextMatrix(lngRow, .ColIndex("缺省价格")) = FormatEx(Val(NVL(rsTemp!单价)), 8)
            .Cell(flexcpData, lngRow, .ColIndex("缺省价格")) = Val(NVL(rsTemp!单价))
            .TextMatrix(lngRow, .ColIndex("缺省执行科室")) = IIF(NVL(rsTemp!执行科室编码) = "", "", NVL(rsTemp!执行科室编码) & "-") & NVL(rsTemp!执行科室名称)
            .Cell(flexcpData, lngRow, .ColIndex("缺省执行科室")) = Val(NVL(rsTemp!执行科室ID))
            .TextMatrix(lngRow, .ColIndex("类别")) = NVL(rsTemp!类别)
            .TextMatrix(lngRow, .ColIndex("是否变价")) = NVL(rsTemp!是否变价)
            .TextMatrix(lngRow, .ColIndex("执行科室")) = NVL(rsTemp!执行科室)
            .TextMatrix(lngRow, .ColIndex("现价")) = IIF(NVL(rsTemp!现价) = "实价", "实价", FormatEx(Val(NVL(rsTemp!现价)), 5))
            .TextMatrix(lngRow, .ColIndex("收费项目")) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = NVL(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("药名")) = ""
            .TextMatrix(lngRow, .ColIndex("药名ID")) = NVL(rsTemp!药名ID)
            .TextMatrix(lngRow, .ColIndex("跟踪在用")) = Val(NVL(rsTemp!跟踪在用))
            If NVL(rsTemp!类别) = "7" Then
                '草药,显示诊疗名称
                .TextMatrix(lngRow, .ColIndex("药名")) = NVL(rsTemp!诊疗编码) & "-" & NVL(rsTemp!诊疗名称)
                .TextMatrix(lngRow, .ColIndex("单位")) = NVL(rsTemp!剂量单位)
                .TextMatrix(lngRow, .ColIndex("缺省数量")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("缺省数量"))) * Val(NVL(rsTemp!剂量系数)), 5)
                .TextMatrix(lngRow, .ColIndex("缺省价格")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("缺省价格"))) / Val(NVL(rsTemp!剂量系数)), 8)
                .TextMatrix(lngRow, .ColIndex("中药形态")) = Val(NVL(rsTemp!中药形态))
                .TextMatrix(lngRow, .ColIndex("剂量系数")) = Val(NVL(rsTemp!剂量系数))
'                If Val(Nvl(rsTemp!中药形态)) = 0 Then   '散装形态的,则显示具体的规格
'                    .TextMatrix(lngRow, .ColIndex("规格")) = Nvl(rsTemp!规格)
'                End If
            Else
                .TextMatrix(lngRow, .ColIndex("单位")) = NVL(rsTemp!计算单位)
            End If
            
            If Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) = 0 Then
                    lng父号 = Val(.Cell(flexcpData, lngRow, .ColIndex("序号")))
                    lng主项ID = Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目")))
                    .IsSubtotal(lngRow) = True: .RowOutlineLevel(lngRow) = 1
            ElseIf lng父号 = Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) Then
                If lngRow > 2 Then
                    If Val(.TextMatrix(lngRow - 1, .ColIndex("从属父号"))) <> 0 Then
                        .IsSubtotal(lngRow - 1) = False
                        .RowOutlineLevel(lngRow - 1) = 2
                    End If
                End If
                .IsSubtotal(lngRow) = True
                .RowOutlineLevel(lngRow) = 2
            End If
            If Not rsOthers Is Nothing And Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) <> 0 Then
                '  "   Select A.主项id, A.从项id, A.固有从属, A.从项数次 "
                rsOthers.Filter = "主项ID=" & lng主项ID & " And 从项ID= " & Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目")))
                If Not rsOthers.EOF Then
                    .TextMatrix(lngRow, .ColIndex("从属数次")) = Val(NVL(rsOthers!从项数次))
                    .Cell(flexcpData, lngRow, .ColIndex("从属父号")) = Val(NVL(rsOthers!固有从属))
                End If
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    '加载使用部门
    lvw科室.ListItems.Clear
    If opt范围(1).value = True Then
         gstrSQL = "" & _
         "   Select A.科室ID,B.编码,b.名称 " & _
         "   From 成套项目使用科室 A,部门表  B  " & _
         "   Where a.科室id=b.Id And a.成套ID=[1]" & _
         "   Order By 编码"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        With lvw科室
             Do While Not rsTemp.EOF
                 .ListItems.Add , "K" & rsTemp!科室ID, NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称), "Dept", "Dept"
                 rsTemp.MoveNext
             Loop
        End With
    End If
    Screen.MousePointer = vbDefault
    Call opt范围_Click(0)
    
    mbln修改 = True
    If opt范围(0).value = True And InStr(mstrPrivs, "修改个人成套方案") < 1 Then
        mbln修改 = False
    ElseIf opt范围(1).value = True And InStr(mstrPrivs, "修改科室成套方案") < 1 Then
        mbln修改 = False
    ElseIf opt范围(2).value = True And InStr(mstrPrivs, "修改全院成套方案") < 1 Then
        mbln修改 = False
    End If
    
    Call EditStatusSet  '设置编辑状态
    ReadCard = True
    Exit Function
ErrHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
End Function

Private Sub cbo人员_Change()
    mblnChange = True
End Sub

Private Sub cbo人员_Click()
    '选择相关的人员
    If cbo人员.ListIndex <> -1 Then
        If cbo人员.ItemData(cbo人员.ListIndex) = 0 Then
             If SearchPerson("") = False Then
             End If
        Else
             cbo人员.Tag = cbo人员.ListIndex
        End If
    End If
End Sub
Private Function SearchPerson(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入指定人员
    '入参:strInput-搜索条件
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-30 16:52:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte, intIdx As Integer
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = gstrLike & strInput & "%"
    strWhere = "": bytStyle = 0
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " And A.编号 Like [1]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = "  And A.简码 Like upper([1])"
        Else
            strWhere = " And (A.编号 Like [1] or A.简码 Like upper([1]) or A.姓名 like [1] )"
        End If
        '2010-12-28 修改(34325)
'        strWhere = strWhere & IIF(gstrNodeNo <> "-", " And (A.站点='" & gstrNodeNo & "' or a.站点 is NULL  )", "")
        
        gstrSQL = "" & _
            "   Select A.ID,A.编号,A.姓名,A.别名,A.简码,A.性别,A.民族,A.出生日期,A.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
            "   From 人员表 A " & _
            "   Where   (A.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & strWhere & _
            "   order by A.编号"
    Else
'        gstrSQL = "" & _
'        "   Select id," & IIF(gstrNodeNo <> "-", "1 as 级数ID,-1*NULL as 上级ID", "Level as 级数ID,上级id") & " ,编码,名称,0 末级,'' as 别名,'' as 简码,''as 性别,''as 民族, to_date(Null,'yyyy-mm-dd')  as 出生日期, '' as  办公室电话 ,'' 执业类别, '' 管理职务,'' 专业技术职务" & _
'        "   From 部门表 " & _
'        "   Where 撤档时间 is null or 撤档时间>=to_date('3000-01-01','yyyy-mm-dd') " & IIF(gstrNodeNo <> "-", " And (A.站点='" & gstrNodeNo & "' or a.站点 is NULL ) ", "") & _
'            IIF(gstrNodeNo <> "-", "", "   Start with 上级id is null connect by prior id=上级id ") & _
'        "   union all " & _
'        "   Select a.ID,999999 AS 级数ID,b.部门id as 上级ID,a.编号,a.姓名,1 as 末级,别名,简码,性别,民族,出生日期,办公室电话,A.执业类别,A.管理职务,A.专业技术职务 " & _
'        "   From 人员表 a,部门人员 b  " & _
'        "   Where a.id=b.人员id and b.缺省=1  " & IIF(gstrNodeNo <> "-", " And (A.站点='" & gstrNodeNo & "' or a.站点 is NULL  )", "") & _
'        "         And (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
'        "   Order by 级数ID,编码"
        
        gstrSQL = "" & _
        "   Select id,1 as 级数ID,-1*NULL as 上级ID,编码,名称,0 末级,'' as 别名,'' as 简码,''as 性别,''as 民族, to_date(Null,'yyyy-mm-dd')  as 出生日期, '' as  办公室电话 ,'' 执业类别, '' 管理职务,'' 专业技术职务" & _
        "   From 部门表 " & _
        "   Where 撤档时间 is null or 撤档时间>=to_date('3000-01-01','yyyy-mm-dd') " & _
        "   union all " & _
        "   Select a.ID,999999 AS 级数ID,b.部门id as 上级ID,a.编号,a.姓名,1 as 末级,别名,简码,性别,民族,出生日期,办公室电话,A.执业类别,A.管理职务,A.专业技术职务 " & _
        "   From 人员表 a,部门人员 b  " & _
        "   Where a.id=b.人员id and b.缺省=1 " & _
        "         And (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        "   Order by 级数ID,编码"
        
        bytStyle = 2
    End If
    
    vRect = zlControl.GetControlRect(cbo人员.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "人员选择", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtParent.Height, blnCancel, False, True, strKey)
    
    
    If blnCancel = True Then
        If cbo人员.Enabled And cbo人员.Visible Then cbo人员.SetFocus
        Call cbo.SetIndex(cbo人员.hwnd, Val(cbo人员.Tag))
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "未找到匹配的使用科室,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If cbo人员.Enabled And cbo人员.Visible Then txtParent.SetFocus
        Call cbo.SetIndex(cbo人员.hwnd, Val(cbo人员.Tag))
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "未找到匹配的使用科室,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If cbo人员.Enabled And cbo人员.Visible Then txtParent.SetFocus
        Call cbo.SetIndex(cbo人员.hwnd, Val(cbo人员.Tag))
        Exit Function
    End If
    If bytStyle = 0 Then
        intIdx = cbo.FindIndex(cbo人员, Val(NVL(rsTemp!ID)))
        If intIdx <> -1 Then
            cbo人员.ListIndex = intIdx
            cbo人员.Tag = cbo人员.ListIndex
        Else
            cbo人员.AddItem rsTemp!编号 & "-" & rsTemp!姓名, 0
            cbo人员.ItemData(cbo人员.NewIndex) = rsTemp!ID
            cbo人员.ListIndex = cbo人员.NewIndex
            cbo人员.Tag = cbo人员.ListIndex
        End If
    Else
        intIdx = cbo.FindIndex(cbo人员, Val(NVL(rsTemp!ID)))
        If intIdx <> -1 Then
            cbo人员.ListIndex = intIdx
            cbo人员.Tag = cbo人员.ListIndex
        Else
            cbo人员.AddItem rsTemp!编码 & "-" & rsTemp!名称, 0
            cbo人员.ItemData(cbo人员.NewIndex) = rsTemp!ID
            cbo人员.ListIndex = cbo人员.NewIndex
            cbo人员.Tag = cbo人员.ListIndex
        End If
    End If
    OS.PressKey vbKeyTab
    SearchPerson = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cbo人员_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(cbo人员.Text) <> "" And cbo人员.ListIndex >= 0 Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If SearchPerson(Trim(cbo人员.Text)) = False Then
        Exit Sub
    End If
End Sub

Private Sub cbo人员_Validate(Cancel As Boolean)
    If cbo人员.ListIndex < 0 Then
        If cbo人员.Text <> "" Then
            Call cbo.SetIndex(cbo人员.hwnd, Val(cbo人员.Tag))
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim ObjItem As ListItem
    If Val(txt科室.Tag) = 0 Then Exit Sub
    If Trim(txt科室.Text) = "" Then Exit Sub
    With lvw科室
        For Each ObjItem In .ListItems
            If Val(Mid(ObjItem.Key, 2)) = Val(txt科室.Tag) Then
                MsgBox "注意:" & vbCrLf & "    该科室已经存在,不能再增加!", vbInformation + vbOKOnly, gstrSysName
                txt科室.SetFocus
                Exit Sub
            End If
        Next
        .ListItems.Add , "K" & txt科室.Tag, txt科室.Text, "Dept", "Dept"
        txt科室.SetFocus
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim intIndex As Integer, strKey As String
    If lvw科室.SelectedItem Is Nothing Then Exit Sub
     With lvw科室
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Call Set使用科室Enable
End Sub
Private Sub Set使用科室Enable()
    '设置使用科室的相关控件状态
    cmdDelete.Enabled = Not Me.lvw科室.SelectedItem Is Nothing
End Sub
Private Sub cmdOK_Click()
    If CheckValied = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mintSucces = mintSucces + 1
    If mEditType = EdI_修改 Then
        Unload Me: mblnChange = False
        Exit Sub
    End If
    Call ClearCardData
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub
Private Function CheckValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据有有效性
    '返回:数据有效,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2010-08-30 15:11:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control, lngRow As Long, blnHaveDate As Boolean
    Dim rsTemp As ADODB.Recordset
    
    If mEditType = EdI_查看 Then Exit Function
        
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")
            If objCtl Is txtCode Or objCtl Is txtName Then
                    If Trim(objCtl.Text) = "" Then
                        MsgBox "注意:" & vbCrLf & "    " & objCtl.Tag & "必须输入,请检查!", vbInformation + vbOKOnly, gstrSysName
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
            End If
            If Not objCtl Is txt科室 Then
                
                If zlStr.ActualLen(Trim(objCtl.Text)) > objCtl.MaxLength Then
                    MsgBox "注意:" & vbCrLf & "    " & objCtl.Tag & "最多能输入" & objCtl.MaxLength & "个字符或" & objCtl.MaxLength \ 2 & "个汉字,请检查!", vbInformation + vbOKOnly, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
                If InStr(1, Trim(objCtl.Text), "'") > 0 Then
                    MsgBox "注意:" & vbCrLf & "    " & objCtl.Tag & "含有非法字符(单引号),请检查!", vbInformation + vbOKOnly, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
            End If
        Case Else
        End Select
    Next
    If Val(txtParent.Tag) = 0 Then
        MsgBox "注意:" & vbCrLf & "    未选择分类信息,请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    With vsWholeSet
        blnHaveDate = False
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目"))) <> 0 Then
                If IsCheckValiedPrice(lngRow, Val(.TextMatrix(lngRow, .ColIndex("缺省价格")))) = False Then
                    .Row = lngRow: .Col = .ColIndex("缺省价格")
                    Call .ShowCell(.Row, .Col)
                    tbPage.Item(0).Selected = True
                    If vsWholeSet.Enabled And vsWholeSet.Visible Then vsWholeSet.SetFocus
                    Exit Function
                End If
                blnHaveDate = True
            End If
        Next
    End With
    If blnHaveDate = False Then
        MsgBox "注意:" & vbCrLf & "    未输入成套收费项目的组成项目,请检查!", vbInformation + vbOKOnly, gstrSysName
        tbPage.Item(0).Selected = True
        If vsWholeSet.Enabled And vsWholeSet.Visible Then vsWholeSet.SetFocus
        Exit Function
    End If
    '检查是否输入了使用科室的
    If opt范围(1).value Then
        If lvw科室.ListItems.Count = 0 Then
            MsgBox "注意:" & vbCrLf & "    未指定使用科室,请检查!", vbInformation + vbOKOnly, gstrSysName
            If tbPage.Item(1).Visible = False Then Exit Function
            tbPage.Item(1).Selected = True
            If txt科室.Enabled And txt科室.Visible Then txt科室.SetFocus
            Exit Function
        End If
    End If
    If opt范围(0).value Then
        If cbo人员.ListIndex < 0 Then
            MsgBox "注意:" & vbCrLf & "    未指定使用人员,请检查!", vbInformation + vbOKOnly, gstrSysName
            If cbo人员.Enabled And cbo人员.Visible Then cbo人员.SetFocus
            Exit Function
        End If
    End If
        
    '检查编码和名称是否重复
    If mEditType = EdI_增加 Or (mEditType = EdI_修改 And Trim(txtCode.Text) <> txtCode.Tag) Then
        gstrSQL = "Select 1 From 成套收费项目 Where 编码 = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtCode.Text))
        If rsTemp.RecordCount > 0 Then
            MsgBox "注意:" & vbCrLf & "    编码已使用，请检查!", vbInformation + vbOKOnly, gstrSysName
            If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
            Exit Function
        End If
    End If
    
    If mEditType = EdI_增加 Or (mEditType = EdI_修改 And Trim(txtName.Text) <> txtName.Tag) Then
        gstrSQL = "Select 1 From 成套收费项目 Where 名称 = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtName.Text))
        If rsTemp.RecordCount > 0 Then
            MsgBox "注意:" & vbCrLf & "    名称已使用，请检查!", vbInformation + vbOKOnly, gstrSysName
            If txtName.Enabled And txtName.Visible Then txtName.SetFocus
            Exit Function
        End If
    End If
    
    CheckValied = True
End Function

Private Sub cmdSel_Click()
    '选择指定的使用部门
    If SearchUseDept("") = False Then Exit Sub
    If cmdAdd.Enabled And cmdAdd.Visible Then cmdAdd.SetFocus
End Sub

Private Sub cmdSelect_Click()
    '选择分类
    If SearchPreLevel("") = False Then Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If ReadCard = False Then Unload Me: Exit Sub
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub

Private Sub Form_Load()
    Call GetPriceGrade(gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    Call InitDefaultLen '取默认字段长度
    Call zlInitClassPage
    Call InitData
    RestoreWinState Me, App.ProductName
    zl_vsGrid_Para_Restore mlngModule, vsWholeSet, Me.Caption, "成套项目组成表列", True, True
    mblnFirst = True
End Sub
Private Function InitData() As Boolean
    Dim strSQL As String
    '执行部门
    On Error GoTo ErrHandle
    strSQL = _
    "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
    " From 部门表 A,部门性质说明 B " & _
    " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
    " And B.部门ID=A.ID and B.服务对象 IN(2,3) " & _
    " Order by B.服务对象,A.编码"
    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo人员.AddItem "[选择人员...]"
    cbo人员.Tag = cbo人员.ListIndex
    InitData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With fra成套
        .Left = ScaleLeft + 50: .Top = ScaleTop + 100
        .Width = ScaleWidth - 100
        tbPage.Top = .Top + .Height + 100
        tbPage.Left = .Left: tbPage.Width = .Width
        picCmd.Top = ScaleHeight - picCmd.Height
        picCmd.Width = ScaleWidth
    End With
    With tbPage
        .Height = picCmd.Top - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Err = 0: On Error Resume Next
  SaveWinState Me, App.ProductName
  Call zlDatabase.SetPara("上次成套方案分类", txtParent.Tag, glngSys, mlngModule, True)
  zl_vsGrid_Para_Save mlngModule, vsWholeSet, Me.Caption, "成套项目组成表列", True, True
End Sub

Private Sub lvw科室_GotFocus()
    Call Set使用科室Enable
End Sub

Private Sub opt范围_Click(Index As Integer)
    mblnChange = True
    Call Set范围状态
End Sub
Private Sub Set范围状态()
    Dim i As Long
    For i = 0 To opt范围.UBound
        If opt范围(i).value Then
            Exit For
        End If
    Next
    Select Case i
    Case 0  '指定人员
        cbo人员.Enabled = True
        cbo人员.BackColor = &H80000005
        If Val(tbPage.Selected.Tag) = mItemPage.pg_使用科室 Then
            tbPage.Item(0).Selected = True
        End If
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = mItemPage.pg_使用科室 Then
                tbPage.Item(i).Visible = False
            End If
        Next
        If cbo人员.ListCount <= 0 Then
            cbo人员.Clear
            cbo人员.AddItem gstrUserName
            cbo人员.ItemData(cbo人员.NewIndex) = glngUserId
            cbo人员.ListIndex = cbo人员.NewIndex
            cbo人员.AddItem "选择其他人员..."
        End If
    Case 1  '指定科室
        cbo人员.Enabled = False
        cbo人员.BackColor = &H8000000F
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = mItemPage.pg_使用科室 Then
                tbPage.Item(i).Visible = True
            End If
        Next
    Case Else
        If Val(tbPage.Selected.Tag) = mItemPage.pg_使用科室 Then
            tbPage.Item(0).Selected = True
        End If
        cbo人员.Enabled = False
        cbo人员.BackColor = &H8000000F
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = mItemPage.pg_使用科室 Then
                tbPage.Item(i).Visible = False
            End If
        Next
    End Select
End Sub
Private Sub opt范围_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub picCmd_Resize()
    Err = 0: On Error Resume Next
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width - 100
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
End Sub
Private Sub picUserDept_Resize()
    Err = 0: On Error Resume Next
    With picUserDept
        lvw科室.Left = 50
        lvw科室.Width = .ScaleWidth - lvw科室.Left * 2
        lvw科室.Height = .ScaleHeight - .Top - 50
    End With
End Sub

 
Private Sub picWholeSet_Resize()
    Err = 0: On Error Resume Next
    With picWholeSet
        vsWholeSet.Left = 50
        vsWholeSet.Width = .ScaleWidth - vsWholeSet.Left * 2
        vsWholeSet.Top = 50
        vsWholeSet.Height = .ScaleHeight - .Top - 50
    End With
End Sub

 
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Val(Item.Tag) = mItemPage.pg_使用科室 Then
        Call Set使用科室Enable
        If txt科室.Enabled And txt科室.Visible Then txt科室.SetFocus
    Else
        If vsWholeSet.Enabled And vsWholeSet.Visible Then vsWholeSet.SetFocus
    End If
End Sub

Private Sub txtCode_Change()
    mblnChange = True
End Sub

Private Sub txtCode_GotFocus()
    zlControl.TxtSelAll txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtCode, KeyAscii, m数字式
End Sub

Private Sub txtMemo_Change()
    mblnChange = True
End Sub

Private Sub txtMemo_GotFocus()
    zlControl.TxtSelAll txtMemo
    OS.OpenIme True
End Sub

Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtMemo, KeyAscii, m文本式
End Sub

Private Sub txtMemo_LostFocus()
    OS.OpenIme False
End Sub

Private Sub txtName_Change()
    mblnChange = True
    txtSymbol.Text = zlStr.GetCodeByORCL(txtName, False, 20)
    txtWB.Text = zlStr.GetCodeByORCL(txtName, True, 20)
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
    OS.OpenIme True
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtName, KeyAscii, m文本式
End Sub

Private Sub txtName_LostFocus()
    OS.OpenIme False
End Sub

Private Sub txtParent_Change()
    mblnChange = True
End Sub

Private Sub txtParent_GotFocus()
    zlControl.TxtSelAll txtParent
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtParent, KeyAscii, m文本式

End Sub

Private Sub txtSymbol_Change()
    mblnChange = True
End Sub

Private Sub txtSymbol_GotFocus()
    zlControl.TxtSelAll txtSymbol
    OS.OpenIme False
End Sub

Private Sub txtSymbol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtSymbol, KeyAscii, m文本式
End Sub

Private Sub txtWB_Change()
    mblnChange = True
End Sub

Private Sub txtWB_GotFocus()
    zlControl.TxtSelAll txtWB
    OS.OpenIme False
End Sub

Private Sub txtWB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtWB_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtWB, KeyAscii, m文本式
End Sub
Private Sub txt科室_Change()
    mblnChange = True: txt科室.Tag = ""
End Sub

Private Sub txt科室_GotFocus()
    zlControl.TxtSelAll txt科室
    Call Set使用科室Enable
End Sub
Private Sub txt科室_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txt科室.Tag) <> "" Then Exit Sub
        If SearchUseDept(Trim(txt科室.Text)) = False Then Exit Sub
        If cmdAdd.Enabled And cmdAdd.Visible Then cmdAdd.SetFocus
End Sub

Private Sub txt科室_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt科室, KeyAscii, m文本式
End Sub
Private Sub ReCale从属项目(ByVal lngRow As Long, ByVal dblNum As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算从属项目数量
    '入参:dblNum-数量
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-31 11:30:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int固定从属 As Integer, i As Long, dblTemp As Double
    With vsWholeSet
        If Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) <> 0 Then Exit Sub
        For i = lngRow + 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("从属父号"))) = Val(.Cell(flexcpData, lngRow, .ColIndex("序号"))) Then
                int固定从属 = Val(.Cell(flexcpData, i, .ColIndex("从属父号")))
                If int固定从属 = 0 Then '非固有从属
                    dblTemp = IIF(dblNum < 0, -1, 1) * Val(.TextMatrix(i, .ColIndex("从属数次")))
                    .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(dblTemp, 5)
                ElseIf int固定从属 = 1 Then '固定的从属
                    dblTemp = IIF(dblNum < 0, -1, 1) * IIF(Val(.TextMatrix(i, .ColIndex("从属数次"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("从属数次"))))
                    .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(dblTemp, 5)
                ElseIf int固定从属 = 2 Then '按比例从属
                    dblTemp = dblNum * Val(.TextMatrix(i, .ColIndex("从属数次")))
                    .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(dblTemp, 5)
                End If
             End If
        Next
    End With
End Sub
Private Sub vsWholeSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关的格式
    '编制:刘兴洪
    '日期:2010-08-27 14:12:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsWholeSet
        Select Case Col
        Case .ColIndex("收费项目")
             .ColComboList(Col) = "..."
        Case .ColIndex("缺省执行科室")
             .ColComboList(Col) = "..."
        Case .ColIndex("缺省数量")
            Call ReCale从属项目(Row, Val(.TextMatrix(Row, .Col)))
        Case .ColIndex("缺省付数")
        Case .ColIndex("缺省价格")
            
        End Select
    End With
End Sub
Private Function IsCheckValiedPrice(ByVal lngRow As Long, ByVal dbl价格 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的价格是否在有效范围内的价格
    '入参:dbl价格
    '出参:
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-16 15:50:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng项目id As Long, dbl最低限价 As Double, dbl最高限价 As Double, dbl缺省价格 As Double
    Dim strMsg As String, rsTemp As ADODB.Recordset
    Dim strSQL  As String, strWherePriceGrade As String
    
    On Error GoTo ErrHandle
    With vsWholeSet
        If dbl价格 = 0 Then '为零,直接退出,不检查
            IsCheckValiedPrice = True: Exit Function
        End If
        lng项目id = Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目")))
        If lng项目id = 0 Then
            .TextMatrix(lngRow, .ColIndex("最低限价")) = ""
            .TextMatrix(lngRow, .ColIndex("最高限价")) = ""
            IsCheckValiedPrice = False
            Exit Function
        End If
        If InStr(1, "5,6,7", Trim(.TextMatrix(lngRow, .ColIndex("类别")))) > 0 Then '药品不检查
            IsCheckValiedPrice = True
            Exit Function
        End If
        If Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "4" And Val(.TextMatrix(lngRow, .ColIndex("跟踪在用"))) = 1 Then '跟踪卫材不检查
            IsCheckValiedPrice = True
            Exit Function
        End If
        
        dbl最低限价 = Val(.TextMatrix(lngRow, .ColIndex("最低限价")))
        dbl最高限价 = Val(.TextMatrix(lngRow, .ColIndex("最高限价")))
        If dbl最低限价 <> 0 And dbl最高限价 <> 0 Then   '已经存在了,则直接返回
            strMsg = CheckScope(dbl最低限价, dbl最高限价, dbl价格)
            If strMsg <> "" Then
                MsgBox "注意:" & vbCrLf & strMsg, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            IsCheckValiedPrice = True: Exit Function
        End If
    End With
    
    If gstr普通价格等级 = "" Or Trim(vsWholeSet.TextMatrix(lngRow, vsWholeSet.ColIndex("类别"))) = "4" Then
        strWherePriceGrade = " And b.价格等级 Is Null"
    Else
        strWherePriceGrade = "" & _
        " And (b.价格等级 = [2]" & vbNewLine & _
        "    Or (b.价格等级 Is Null" & vbNewLine & _
        "        And Not Exists(Select 1" & vbNewLine & _
        "                       From 收费价目" & vbNewLine & _
        "                       Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
        "                             And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    End If
    strSQL = _
          " Select  B.现价,B.原价,B.缺省价格 " & _
          " From  收费价目 B" & _
          " Where  Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
          "        And B.收费细目ID=[1]" & strWherePriceGrade
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, gstr普通价格等级)
    If rsTemp.EOF = False Then
        dbl最低限价 = Val(NVL(rsTemp!原价))
        dbl最高限价 = Val(NVL(rsTemp!现价))
        dbl缺省价格 = Val(NVL(rsTemp!缺省价格))
        With vsWholeSet
            .TextMatrix(lngRow, .ColIndex("最低限价")) = dbl最低限价
            .TextMatrix(lngRow, .ColIndex("最高限价")) = dbl最高限价
            .Cell(flexcpData, lngRow, .ColIndex("最高限价")) = dbl缺省价格
        End With
        strMsg = CheckScope(dbl最低限价, dbl最高限价, dbl价格)
        If strMsg <> "" Then
            MsgBox "注意:" & vbCrLf & strMsg, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    IsCheckValiedPrice = True: Exit Function
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsWholeSet_AfterMoveColumn(ByVal Col As Long, Position As Long)
  zl_vsGrid_Para_Save mlngModule, vsWholeSet, Me.Caption, "成套项目组成表列", True, True
End Sub

Private Sub vsWholeSet_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnSort = True Then Exit Sub
    Call zl_VsGridRowChange(vsWholeSet, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Function IsHaveHypotaxisItem(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该项是否包含从属项目
    '返回:包含返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-31 12:00:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsWholeSet
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("从属父号"))) = lngRow Then
                IsHaveHypotaxisItem = True: Exit Function
            End If
        Next
    End With
End Function

Private Sub vsWholeSet_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("标志")
        Case Else
        End Select
    End With
  zl_vsGrid_Para_Save mlngModule, vsWholeSet, Me.Caption, "成套项目组成表列", True, True

End Sub

Private Sub vsWholeSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, arrSplit As Variant
    With vsWholeSet
        If mEditType = EdI_查看 Then
            Cancel = True: Exit Sub
        End If
        Select Case Col
        Case .ColIndex("收费项目")
            If Val(.TextMatrix(Row, .ColIndex("从属父号"))) <> 0 Then
                Cancel = True: Exit Sub
            ElseIf IsHaveHypotaxisItem(Row) Then
                Cancel = True
            End If
        Case .ColIndex("缺省执行科室")
            If Not IsEdit执行科室 Then Cancel = True
        Case .ColIndex("缺省数量")
        Case .ColIndex("缺省付数")
            If InStr(1, ",7", Trim(.TextMatrix(Row, .ColIndex("类别")))) = 0 Then Cancel = True: Exit Sub
        Case .ColIndex("缺省价格")
            If Val(.TextMatrix(Row, .ColIndex("是否变价"))) = 0 Then Cancel = True: Exit Sub
            '药品和跟踪在用的卫生材料的价格是重新算的,所以不用设置缺省价格
            If InStr(1, "5,6,7", Trim(.TextMatrix(Row, .ColIndex("类别")))) > 0 Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("类别"))) = "4" And Val(.TextMatrix(Row, .ColIndex("跟踪在用"))) = 1 Then Cancel = True: Exit Sub
        Case Else: Cancel = True
        End Select
    End With
End Sub

Private Sub vsWholeSet_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("标志")
            Cancel = True
            
        Case Else
        End Select
    End With

End Sub

Private Sub vsWholeSet_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long
    If mEditType = EdI_查看 Then Exit Sub
    
    With vsWholeSet
        Select Case Col
        Case .ColIndex("收费项目")
            If Select收费项目("") = False Then Exit Sub
            Call zlVsMoveGridCell(vsWholeSet, .ColIndex("收费项目"), , True, lngRow)
        Case .ColIndex("缺省执行科室")
            If ShowSelectDept("") = False Then Exit Sub
            Call zlVsMoveGridCell(vsWholeSet, .ColIndex("收费项目"), , True, lngRow)
        End Select
    End With
End Sub

Private Sub vsWholeSet_CellChanged(ByVal Row As Long, ByVal Col As Long)
  mblnChange = True
End Sub
Private Sub SetInputFormat(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入格式设置
    '编制:刘兴洪
    '日期:2010-08-27 14:42:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrSplit As Variant
    If mEditType = EdI_查看 Then Exit Sub
    With vsWholeSet
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("规格")) = "1||1"
        .ColData(.ColIndex("序号")) = "1||1"
        .ColData(.ColIndex("单位")) = "0||1"
        .ColData(.ColIndex("从属父号")) = "1||1"
        .ColData(.ColIndex("是否变价")) = "1||1"
        .ColData(.ColIndex("类别")) = "1||1"
        .ColData(.ColIndex("执行科室")) = "1||1"
        .ColData(.ColIndex("从属数次")) = "1||1"
    End With
End Sub

Private Sub vsWholeSet_EnterCell()
    If mblnSort = True Then Exit Sub
    If mEditType = EdI_查看 Then Exit Sub
    
    '新增或修改才存在设置
    With vsWholeSet
        SetInputFormat .Row
        OS.OpenIme (False)
        Select Case .Col
        Case .ColIndex("收费项目")
             .ColComboList(.Col) = "..."
        Case .ColIndex("缺省执行科室")
             .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsWholeSet_GotFocus()
  Call zl_VsGridGotFocus(vsWholeSet)
End Sub

Private Sub vsWholeSet_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    Dim blnDeleteSubs As Boolean  '是否删除从属项目
    Dim blnHaveData As Boolean, lng父号 As Long
    With vsWholeSet
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyReturn _
            And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                vsWholeSet_CellButtonClick .Row, .Col
            Else
            Select Case .Col
            Case .ColIndex("收费项目")
                .ColComboList(.Col) = ""
            Case .ColIndex("缺省执行科室")
                .ColComboList(.Col) = ""
            Case Else
            End Select
            End If
        End If
 
        If KeyCode = vbKeyDelete Then
            blnCancel = False
            '删除行前
            Call BeforeDeleteRow(.Row, blnCancel, blnDeleteSubs)
            If blnCancel = True Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                If Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) = 0 Then
                    Do While True
                        blnHaveData = False
                         For lngRow = .Row + 1 To .Rows - 1
                            If Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) = Val(.Cell(flexcpData, .Row, .ColIndex("序号"))) Then
                                If blnDeleteSubs Then
                                    .RemoveItem lngRow
                                    blnHaveData = True
                                    Exit For
                                Else
                                    .TextMatrix(lngRow, .ColIndex("从属父号")) = ""
                                    .IsSubtotal(lngRow) = True
                                    .RowOutlineLevel(lngRow) = 1
                                End If
                            End If
                        Next
                        If blnHaveData = False Then Exit Do
                    Loop
                    If .Row = -1 Then .Row = .Rows - 1
                End If
                If .Row = .Rows - 1 And .Row = 1 Then
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(.Row, lngCol) = ""
                        .Cell(flexcpData, .Row, lngCol) = ""
                    Next
                    .IsSubtotal(.Row) = True
                    .RowOutlineLevel(.Row) = 1
                Else
                    .RemoveItem .Row
                End If
            End If
            '删除行后
            Call AfterDeleteRow
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsWholeSet
        If Trim(.TextMatrix(.Row, .ColIndex("收费项目"))) = "" Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsWholeSet, .ColIndex("收费项目"), , IIF(mEditType <> EdI_查看, True, False), lngRow)
        If lngRow >= 0 Then
            Call AfterAddRow(lngRow)
        End If
    End With
End Sub
Private Sub AfterAddRow(Row As Long)
    '增加行后
    Call RefreshRowNO(Row)
End Sub
Private Sub AfterDeleteRow()
    '删除行后
    Call RefreshRowNO
End Sub
Private Sub RefreshRowNO(Optional lngRow As Long = 1)
    Dim i As Long, j As Long, lng序号 As Long
    '重新计算序号
    With vsWholeSet
        '重新算序号
        For i = lngRow To .Rows - 1
            .TextMatrix(i, .ColIndex("序号")) = i
            lng序号 = Val(.Cell(flexcpData, i, .ColIndex("序号")))
            For j = i + 1 To .Rows - 1
                If Val(.TextMatrix(j, .ColIndex("从属父号"))) = lng序号 And lng序号 <> 0 Then
                    .TextMatrix(j, .ColIndex("从属父号")) = i
                End If
            Next
            .Cell(flexcpData, i, .ColIndex("序号")) = i
            If Trim(.TextMatrix(i, .ColIndex("收费项目"))) = "" Then
                .IsSubtotal(i) = True
                .RowOutlineLevel(i) = 1
            End If
        Next
    End With
End Sub

Private Sub vsWholeSet_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsWholeSet
        Select Case Col
        Case .ColIndex("收费项目")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If Select收费项目(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
        Case .ColIndex("缺省执行科室")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If ShowSelectDept(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vsWholeSet, .ColIndex("收费项目"), -1, True, lngRow)
        If lngRow >= 0 Then AfterAddRow lngRow
    End With
End Sub

Private Sub vsWholeSet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsWholeSet_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsWholeSet
        Select Case .Col
            Case .ColIndex("收费项目")
                Grid.CheckKeyPress vsWholeSet, Row, Col, KeyAscii, m文本式
            Case .ColIndex("缺省数量"), .ColIndex("缺省价格"), .ColIndex("缺省付数")
                Grid.CheckKeyPress vsWholeSet, Row, Col, KeyAscii, m金额式
        End Select
    End With
End Sub

Private Sub vsWholeSet_LeaveCell()
    If mblnSort Then Exit Sub
    OS.OpenIme False
End Sub

Private Sub vsWholeSet_LostFocus()
    OS.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsWholeSet)
End Sub

Private Sub vsWholeSet_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '设置单元格的编辑长度
        With vsWholeSet
           Select Case .Col
               Case .ColIndex("收费项目")
                   .EditMaxLength = 100
               Case .ColIndex("缺省数量"), .ColIndex("缺省价格")
                   .EditMaxLength = 16
               Case .ColIndex("缺省付数")
                  .EditMaxLength = 3
                  ' .EditMask = "-1234567890"
           End Select
    End With
End Sub

Private Sub vsWholeSet_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String
    '数据验证
    With vsWholeSet
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("缺省数量")
                If zlNumInputCheck(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = FormatEx(Val(strKey), 5)
                End If
            Case .ColIndex("缺省付数")
                If zlNumInputCheck(strKey, 3, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = IIF(Val(strKey) = 0, 1, Val(strKey))
                End If
            Case .ColIndex("缺省价格")
                If zlNumInputCheck(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If IsCheckValiedPrice(Row, Val(strKey)) = False Then
                     Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = FormatEx(Val(strKey), 8)
                End If
        End Select
    End With
End Sub
Private Sub BeforeDeleteRow(Row As Long, Cancel As Boolean, blnDeleteSubs As Boolean)
    If mEditType = EdI_查看 Then Cancel = True: Exit Sub
    With vsWholeSet
        If Val(.Cell(flexcpData, Row, .ColIndex("收费项目"))) <> 0 Then
            If MsgBox("你是否真的要删除收费项目为“" & .TextMatrix(.Row, .ColIndex("收费项目")) & "”的记录吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            If IsHaveHypotaxisItem(Row) Then
                If MsgBox("收费项目为“" & .TextMatrix(.Row, .ColIndex("收费项目")) & "”有从属项目,是否连同从属项目一并删除?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    blnDeleteSubs = True
                End If
            End If
            If Val(.TextMatrix(Row, .ColIndex("从属父号"))) <> 0 And .IsSubtotal(Row) Then
                '删除的是从属父号
                '看上一条是否从属序号一致
                If Row >= 2 Then
                    If .TextMatrix(Row - 1, .ColIndex("从属父号")) = .TextMatrix(Row, .ColIndex("从属父号")) Then
                        .IsSubtotal(Row - 1) = True
                        .RowOutlineLevel(Row - 1) = 2
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function Select收费项目(Optional strSearch As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费项目选择器
    '入参:strSearch-要搜索的条件(""表示弹出所有卫材)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-27 14:38:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, sngLeft As Single, sngTop As Single
    Dim int病人来源  As Integer, str类别 As String, lng项目id As Long, str排除类别 As String
    Dim j As Long
    
    With vsWholeSet
        str排除类别 = ""
        If .TextMatrix(1, .ColIndex("类别")) <> "7" And .TextMatrix(1, .ColIndex("类别")) <> "" Then
            str排除类别 = "'7'"
        End If
        If .TextMatrix(1, .ColIndex("类别")) = "7" Then
            str类别 = "'7'"
        End If
        
        int病人来源 = 2
        If strSearch = "" Then
            lng项目id = frmItemSelect.ShowSelect(Me, int病人来源, True, str类别, , , "", zl获取中药形态(.Row), str排除类别, , gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
        Else
            'Call CalcPosition(sngLeft, sngTop, vsWholeSet)
            lng项目id = frmItemSelect.ShowSelect(Me, int病人来源, True, str类别, strSearch, .EditWindow, "", zl获取中药形态(.Row), str排除类别, , gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
        End If
        If lng项目id = 0 Then GoTo NotSel:
        
        If CheckItemsIsExsits(lng项目id, .Row) Then GoTo NotSel:
        
        If LoadWholeItem(lng项目id) = False Then
            GoTo NotSel:
        End If
        Select收费项目 = True
NotSel:
        vsWholeSet_GotFocus
    End With
End Function
Private Function CheckItemsIsExsits(ByVal lng项目id As Long, ByVal lngNotCheckRow As Long, Optional blnSubItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查收费细目是否存在
    '入参:lngNotCheckRow-不检查的行
    '       blnSubItem-是否套参项目
    '出参:
    '返回:存在项目,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-31 16:00:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    CheckItemsIsExsits = True
    With vsWholeSet
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目"))) = lng项目id And lngNotCheckRow <> lngRow Then
                If blnSubItem = True Then
                    MsgBox "注意:" & vbCrLf & "   在第" & lngRow & "行中已经存在" & vbCrLf & "  〖" & .TextMatrix(lngRow, .ColIndex("收费项目")) & " 〗 " & vbCrLf & "项目了,不能加载该从属项!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                Else
                    MsgBox "在第" & lngRow & "行中已经存在该项目了,请检查!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    CheckItemsIsExsits = False
End Function


Public Function zl获取中药形态(Optional ByVal lngRow As Long = -1, Optional blnOnly中成药 As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据是否录入了中草药的
    '入参:blnOnly中成药-仅判断是否有中成药(对配方时判断有效):原因是中成药在配方中已经存在,就不需要检查
    '     lngRow-当前操作的行
    '出参:
    '返回:录入了中草药的,则返回中药形态属性(0-散装,1-饮片,2-免煎剂),否则返回-1 表示还没有录入中药形态项目
    '编制:刘兴洪
    '日期:2010-02-02 11:44:17
    '问题:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    zl获取中药形态 = -1
    '如果未指定页,则用当前页
    strTemp = IIF(blnOnly中成药, ",6,", ",6,7,")
    With vsWholeSet
        For i = 1 To .Rows - 1
            If InStr(1, strTemp, "," & .TextMatrix(i, .ColIndex("类别")) & ",") > 0 And Val(.Cell(flexcpData, i, .ColIndex("收费项目"))) <> 0 And i <> lngRow Then
                zl获取中药形态 = Val(.TextMatrix(i, .ColIndex("中药形态")))
                Exit Function
            End If
        Next
    End With
End Function
Private Function zlNumInputCheck(ByVal strInput As String, ByVal intMax As Integer, Optional bln负数检查 As Boolean = True, Optional bln零检查 As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的数量或金额
    '入参::strInput        输入的字符串
    '     intMax          整数的位数
    '     bln负数检查     是否进行负数检查
    '     bln零检查         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-27 15:03:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblValue As Double
    If bln零检查 = True Then
        If strInput = "" Then
            ShowMsgBox str项目 & "未输入，请检查!"
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlNumInputCheck = True: Exit Function
    
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If bln负数检查 = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    
    If bln零检查 = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    zlNumInputCheck = True
End Function
Private Function SearchPreLevel(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择上级分类
    '返回:
    '编制:刘兴洪
    '日期:2010-08-26 13:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = gstrLike & strInput & "%"
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " 编码 Like [1]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = " 简码 Like upper([1])"
        Else
            strWhere = " 编码 Like [1] or 简码 Like upper([1]) or 名称 like [1]"
        End If
        gstrSQL = "" & _
        " Select ID,上级ID,编码,名称,简码" & _
        " From 成套项目分类" & _
        " Where " & strWhere
        bytStyle = 0
    Else
        gstrSQL = "" & _
        " Select ID,上级ID,编码,名称,简码" & _
        " From 成套项目分类" & _
        " Start with 上级ID is null Connect by prior ID=上级ID"
        bytStyle = 1
    End If
    
    vRect = zlControl.GetControlRect(txtParent.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "成套收费项目分类", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtParent.Height, blnCancel, False, True, strKey)
    
    If blnCancel = True Then
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "未找到匹配的分类信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "未找到匹配的分类信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    txtParent.Text = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
    txtParent.Tag = NVL(rsTemp!ID)
    Call zlDefaultCode
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    SearchPreLevel = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SearchUseDept(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择使用科室
    '返回:
    '编制:刘兴洪
    '日期:2010-08-26 13:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte, str范围 As String
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = gstrLike & strInput & "%"
    strWhere = ""
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " And A.编码 Like [3]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = "  And A.简码 Like upper([3])"
        Else
            strWhere = " And (A.编码 Like [3] or A.简码 Like upper([3]) or A.名称 like [3] )"
        End If
    End If
    
    str范围 = "1,2,3"
    'str范围 = IIF(chk范围(0).Value = 1, ",1", "") & IIF(chk范围(1).Value = 1, ",2", "") & ",3,"
    If InStr(1, mstrPrivs, ";本院成套方案;") > 0 Then
        '可以指定的全院科室
        gstrSQL = "" & _
        "   Select Distinct A.ID,A.名称,A.编码 " & _
        "   From 部门表 A,部门性质说明 B" & _
        "   Where A.ID=B.部门ID  " & strWhere & _
        "        And Instr([1],B.服务对象)>0" & _
        "       And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is Null)" & _
        "       And B.工作性质 IN('临床','护理','检查','检验','手术','治疗','营养')" & _
        " Order by A.编码"
    Else
        '只能指定自已的科室
        gstrSQL = "" & _
            "   Select Distinct A.ID,A.名称,A.编码  " & _
            "   From 部门表 A,部门性质说明 B,部门人员 C" & _
            "   Where A.ID=B.部门ID " & strWhere & _
            "       And Instr([1],B.服务对象)>0 And A.ID=C.部门ID And C.人员ID=[2]" & _
            "       And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is Null)" & _
            "       And B.工作性质 IN('临床','护理','检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    End If
    
    vRect = zlControl.GetControlRect(txt科室.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "使用科室选择", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txt科室.Height, blnCancel, False, True, str范围, glngUserId, strKey)
    
    If blnCancel = True Then
        If txt科室.Enabled And txt科室.Visible Then txt科室.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "未找到匹配的使用科室,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txt科室.Enabled And txt科室.Visible Then txt科室.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "未找到匹配的使用科室,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txt科室.Enabled And txt科室.Visible Then txt科室.SetFocus
        Exit Function
    End If
    txt科室.Text = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
    txt科室.Tag = NVL(rsTemp!ID)
    Call Set使用科室Enable
    SearchUseDept = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存卡片数据信息
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-30 15:35:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, cllPro As Collection, lngId As Long
    Dim ObjItem As ListItem, dblTemp As Double
    On Error GoTo ErrHandle
    Set cllPro = New Collection
    If mEditType = EdI_增加 Then
        lngId = Sys.NextId("成套收费项目")
        'Zl_成套收费项目_Insert
        gstrSQL = "Zl_成套收费项目_Insert("
    Else
        lngId = Val(mstrID)
        '需要删除相关的成套项目组成和使用科室
        'Zl_成套收费项目_Update
        gstrSQL = "Zl_成套收费项目_Update("
    End If
    '  Id_In     In 成套收费项目.ID%Type,
    gstrSQL = gstrSQL & "" & lngId & ","
    '  分类id_In In 成套收费项目.分类id%Type,
    gstrSQL = gstrSQL & "" & IIF(Val(txtParent.Tag) = 0, "NULL", Val(txtParent.Tag)) & ","
    '  编码_In   In 成套收费项目.编码%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtCode.Text) & "',"
    '  名称_In   In 成套收费项目.名称%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtName.Text) & "',"
    '  范围_In   In 成套收费项目.范围%Type,
    gstrSQL = gstrSQL & "" & IIF(opt范围(0).value, 2, IIF(opt范围(1).value, 1, 0)) & ","
    '  人员id_In In 成套收费项目.人员id%Type,
    If opt范围(0).value Then
        gstrSQL = gstrSQL & "" & cbo人员.ItemData(cbo人员.ListIndex) & ","
    Else
        gstrSQL = gstrSQL & "NULL" & ","
    End If
    '  五笔_In   In 成套收费项目.五笔%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtWB.Text) & "',"
    '  备注_In   In 成套收费项目.备注%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtMemo.Text) & "',"
    '  拼音_In   In 成套收费项目.拼音%Type
    gstrSQL = gstrSQL & "'" & Trim(txtSymbol.Text) & "')"
    zlDatabase.AddItem cllPro, gstrSQL
    '加入组成部分
    With vsWholeSet
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目"))) <> 0 Then
                'Zl_成套收费项目组合_Insert
                gstrSQL = "Zl_成套收费项目组合_Insert("
                '  成套id_In     In 成套收费项目组合.成套id%Type,
                gstrSQL = gstrSQL & "" & lngId & ","
                '  收费细目id_In In 成套收费项目组合.收费细目id%Type,
                gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目"))) & ","
                '  序号_In       In 成套收费项目组合.序号%Type,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("序号"))) & ","
                '  从属父号_In   In 成套收费项目组合.从属父号%Type,
                gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(lngRow, .ColIndex("从属父号"))) = 0, "NULL", Val(.TextMatrix(lngRow, .ColIndex("从属父号")))) & ","
                '付数_IN
                dblTemp = Val(.TextMatrix(lngRow, .ColIndex("缺省付数")))
                gstrSQL = gstrSQL & "" & dblTemp & ","
                '  数量_In       In 成套收费项目组合.数量%Type,
                dblTemp = Val(.TextMatrix(lngRow, .ColIndex("缺省数量")))
                If .TextMatrix(lngRow, .ColIndex("类别")) = "7" Then
                    If Val(.TextMatrix(lngRow, .ColIndex("剂量系数"))) <> 0 Then
                        dblTemp = Round(dblTemp / Val(.TextMatrix(lngRow, .ColIndex("剂量系数"))), 5)
                    End If
                End If
                gstrSQL = gstrSQL & "" & dblTemp & ","
                '  单价_In       In 成套收费项目组合.单价%Type,
                dblTemp = Val(.TextMatrix(lngRow, .ColIndex("缺省价格")))
                If .TextMatrix(lngRow, .ColIndex("类别")) = "7" Then
                    If Val(.TextMatrix(lngRow, .ColIndex("剂量系数"))) <> 0 Then
                        dblTemp = Round(dblTemp * Val(.TextMatrix(lngRow, .ColIndex("剂量系数"))), 8)
                    End If
                End If
                gstrSQL = gstrSQL & "" & dblTemp & ","
                '  执行科室id_In In 成套收费项目组合.执行科室id%Type
                If Val(.Cell(flexcpData, lngRow, .ColIndex("缺省执行科室"))) <> 0 Then
                        gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("缺省执行科室"))) & ")"
                Else
                        gstrSQL = gstrSQL & "NULL)"
                End If
                zlDatabase.AddItem cllPro, gstrSQL
            End If
        Next
    End With
    '加入使用科室
    If opt范围(1).value Then
        With lvw科室
                For Each ObjItem In lvw科室.ListItems
                    'Zl_成套项目使用科室_Insert
                    gstrSQL = "Zl_成套项目使用科室_Insert ("
                    '  成套id_In In 成套项目使用科室.成套id%Type,
                    gstrSQL = gstrSQL & "" & lngId & ","
                    '  科室id_In In 成套项目使用科室.科室id%Type
                    gstrSQL = gstrSQL & "" & Val(Mid(ObjItem.Key, 2)) & ")"
                     zlDatabase.AddItem cllPro, gstrSQL
                Next
        End With
    End If
    Err = 0: On Error GoTo ErrCommit:
    zlDatabase.ExecuteProcedureBeach cllPro, Me.Caption
    SaveData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrCommit:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function LoadWholeItem(ByVal lng项目id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载指定的收费细目
    '入参:lng项目ID-收费细目ID
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-30 17:57:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strWherePriceGrade As String
    
    If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
        strWherePriceGrade = " And j.价格等级 Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And j.价格等级 = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And j.价格等级 = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And j.价格等级 = [4])" & vbNewLine & _
            "      Or (j.价格等级 Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From 收费价目" & vbNewLine & _
            "                          Where j.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And 价格等级 = [4])))))"
    End If
    strSQL = _
    " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
    "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,0 as 要求审批," & _
    "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
    "       Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
    "       Decode(A.类别,'4',1,C.住院包装) as 住院包装," & _
    "       Decode(A.类别,'4',A.计算单位,C.住院单位) as 住院单位,D.跟踪在用,A.录入限量,C.中药形态," & _
    "       M1.编码 as 诊疗编码,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,C.剂量系数," & _
    "       Decode(A.是否变价,1,'时价',LTrim(To_Char(J1.现价,'999999999.9999999'))) as 现价" & _
    " From 收费项目目录 A, 收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,诊疗项目目录 M1," & _
    "             (Select j.收费细目id, Sum(j.现价) as 现价" & vbNewLine & _
    "              From 收费价目 J,收费项目目录 K" & vbNewLine & _
    "              Where j.收费细目ID = k.ID And Sysdate Between J.执行日期 And Nvl(J.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.收费细目id ) J1 " & _
    " Where A.ID=J1.收费细目ID(+)  " & _
    "       And A.类别=B.编码 And A.ID=C.药品ID(+) And C.药名ID=M1.id(+) And A.ID=D.材料ID(+)" & _
    "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIF(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "       And A.ID=[1] "
      
    
 On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    If rsTemp.EOF Then Exit Function
    If NVL(rsTemp!类别) = "7" Then
        '需要检查药名是否存在
        If ItemExist(Val(NVL(rsTemp!药名ID)), vsWholeSet.Row) Then
            ShowMsgBox "注意:" & vbCrLf & "   草药名为" & NVL(rsTemp!诊疗名称) & " 已经存在,不能再输入!"
            Exit Function
        End If
    End If
    
    With vsWholeSet
        '当前行:
        .TextMatrix(.Row, .ColIndex("类别")) = NVL(rsTemp!类别)
        .TextMatrix(.Row, .ColIndex("收费项目")) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
        .EditText = .TextMatrix(.Row, .ColIndex("收费项目"))
        .Cell(flexcpData, .Row, .ColIndex("收费项目")) = NVL(rsTemp!ID)
        .TextMatrix(.Row, .ColIndex("规格")) = NVL(rsTemp!规格)
        .TextMatrix(.Row, .ColIndex("中药形态")) = Val(NVL(rsTemp!中药形态))
        .TextMatrix(.Row, .ColIndex("剂量系数")) = Val(NVL(rsTemp!剂量系数))
        .TextMatrix(.Row, .ColIndex("药名ID")) = NVL(rsTemp!药名ID)
        .TextMatrix(.Row, .ColIndex("单位")) = NVL(rsTemp!计算单位)
        .TextMatrix(.Row, .ColIndex("跟踪在用")) = Val(NVL(rsTemp!跟踪在用))
        .TextMatrix(.Row, .ColIndex("最低限价")) = ""
        .TextMatrix(.Row, .ColIndex("最高限价")) = ""
        .TextMatrix(.Row, .ColIndex("现价")) = IIF(NVL(rsTemp!现价) = "实价", "实价", FormatEx(Val(NVL(rsTemp!现价)), 5))
        If Val(.TextMatrix(.Row, .ColIndex("缺省付数"))) = 0 Then
            .TextMatrix(.Row, .ColIndex("缺省付数")) = 1
        End If
        If NVL(rsTemp!类别) = "7" Then
            '草药,显示诊疗名称
            .TextMatrix(.Row, .ColIndex("药名")) = NVL(rsTemp!诊疗编码) & "-" & NVL(rsTemp!诊疗名称)
            .TextMatrix(.Row, .ColIndex("单位")) = NVL(rsTemp!剂量单位)
            .TextMatrix(.Row, .ColIndex("缺省数量")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("缺省数量"))) * Val(NVL(rsTemp!剂量系数)), 5)
            .TextMatrix(.Row, .ColIndex("缺省价格")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("缺省价格"))) / Val(NVL(rsTemp!剂量系数)), 8)
            .TextMatrix(.Row, .ColIndex("中药形态")) = Val(NVL(rsTemp!中药形态))
        End If
        .TextMatrix(.Row, .ColIndex("从属父号")) = ""
        .TextMatrix(.Row, .ColIndex("序号")) = .Row
        .Cell(flexcpData, .Row, .ColIndex("序号")) = .Row
        .TextMatrix(.Row, .ColIndex("是否变价")) = NVL(rsTemp!是否变价)
        .TextMatrix(.Row, .ColIndex("执行科室")) = NVL(rsTemp!执行科室)
        .IsSubtotal(.Row) = True
        .RowOutlineLevel(.Row) = 1
        
        If InStr(",5,6,7,", NVL(rsTemp!类别)) = 0 Then
            '药品不能设置从属项目
            If (gbln从项汇总折扣 And Val(.TextMatrix(.Row, .ColIndex("从属父号"))) = 0) Or Not gbln从项汇总折扣 Then  '(如果有级联,只取一级)
                If CheckISGetSubItem(.Row) Then
                     '加载子项
                     Call LoadWholeSubItems(Val(.Cell(flexcpData, .Row, .ColIndex("收费项目"))), .Row)
                End If
            End If
        End If
        .Cell(flexcpData, .Row, .ColIndex("序号")) = .Row
    End With
    LoadWholeItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadWholeSubItems(ByVal lng项目id As Long, ByVal lng父号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载套餐项目
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-31 10:29:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, i As Long, lngRow As Long
    Dim strWherePriceGrade As String
    
    If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
        strWherePriceGrade = " And j.价格等级 Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And j.价格等级 = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And j.价格等级 = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And j.价格等级 = [4])" & vbNewLine & _
            "      Or (j.价格等级 Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From 收费价目" & vbNewLine & _
            "                          Where j.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And 价格等级 = [4])))))"
    End If
    strSQL = _
    "Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
    "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别,0 as 要求审批," & _
    "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
    "       Decode(A.类别,'4',1,D.住院包装) as 住院包装,A.服务对象," & _
    "       Decode(A.类别,'4',A.计算单位,D.住院单位) as 住院单位," & _
    "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,D.中药形态," & _
    "       M1.编码 as 诊疗编码,M1.名称 as 诊疗名称,M1.计算单位 as 剂量单位,D.剂量系数," & _
    "       Decode(A.是否变价,1,'时价',LTrim(To_Char(J1.现价,'999999999.9999999'))) as 现价" & _
    " From 收费项目目录 A, 收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1,诊疗项目目录 M1," & _
    "             (Select j.收费细目id, Sum(j.现价) as 现价" & vbNewLine & _
    "              From 收费价目 J,收费项目目录 K" & vbNewLine & _
    "              Where j.收费细目ID = k.ID And Sysdate Between J.执行日期 And Nvl(J.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.收费细目id ) J1 " & _
    " Where A.ID=J1.收费细目ID(+)  " & _
    "   And B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And D.药名ID=M1.id(+)  And A.ID=E.材料ID(+)" & _
    "   And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
    "   And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIF(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "   And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "   And C.主项ID=[1] " & _
    " Order by A.编码"

    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    With vsWholeSet
        .RowData(lng父号) = 1
        Do While Not rsTemp.EOF
            If CheckItemsIsExsits(Val(NVL(rsTemp!ID)), 0, True) = False Then
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                '当前行:
                .TextMatrix(lngRow, .ColIndex("序号")) = lngRow
                .Cell(flexcpData, lngRow, .ColIndex("序号")) = lngRow
                .TextMatrix(lngRow, .ColIndex("类别")) = NVL(rsTemp!类别)
                .TextMatrix(lngRow, .ColIndex("从属父号")) = lng父号
                .Cell(flexcpData, lngRow, .ColIndex("从属父号")) = Val(NVL(rsTemp!固有从属))
                .TextMatrix(lngRow, .ColIndex("从属数次")) = Val(NVL(rsTemp!从项数次))
                .TextMatrix(lngRow, .ColIndex("收费项目")) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
                .Cell(flexcpData, lngRow, .ColIndex("收费项目")) = NVL(rsTemp!ID)
                .TextMatrix(lngRow, .ColIndex("规格")) = NVL(rsTemp!规格)
                .TextMatrix(lngRow, .ColIndex("中药形态")) = Val(NVL(rsTemp!中药形态))
                .TextMatrix(lngRow, .ColIndex("单位")) = NVL(rsTemp!计算单位)
                .TextMatrix(lngRow, .ColIndex("跟踪在用")) = Val(NVL(rsTemp!跟踪在用))
                .TextMatrix(lngRow, .ColIndex("剂量系数")) = Val(NVL(rsTemp!剂量系数))
                .TextMatrix(lngRow, .ColIndex("药名ID")) = NVL(rsTemp!药名ID)
                .TextMatrix(lngRow, .ColIndex("最低限价")) = ""
                .TextMatrix(lngRow, .ColIndex("最高限价")) = ""
                .TextMatrix(lngRow, .ColIndex("现价")) = IIF(NVL(rsTemp!现价) = "实价", "实价", FormatEx(Val(NVL(rsTemp!现价)), 5))
                If Val(.TextMatrix(lngRow, .ColIndex("缺省付数"))) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("缺省付数")) = 1
                End If
                If NVL(rsTemp!类别) = "7" Then
                    '草药,显示诊疗名称
                    .TextMatrix(lngRow, .ColIndex("药名")) = NVL(rsTemp!诊疗编码) & "-" & NVL(rsTemp!诊疗名称)
                    .TextMatrix(lngRow, .ColIndex("单位")) = NVL(rsTemp!剂量单位)
                    .TextMatrix(lngRow, .ColIndex("缺省数量")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("缺省数量"))) * Val(NVL(rsTemp!剂量系数)), 5)
                    .TextMatrix(lngRow, .ColIndex("缺省价格")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("缺省价格"))) / Val(NVL(rsTemp!剂量系数)), 8)
                End If
                
                .TextMatrix(lngRow, .ColIndex("是否变价")) = NVL(rsTemp!是否变价)
                .TextMatrix(lngRow, .ColIndex("执行科室")) = NVL(rsTemp!执行科室)
                .RowData(lngRow) = 0
                
                If lng父号 <> 0 Then  '对上级进行分级
                      .IsSubtotal(lng父号) = True: .RowOutlineLevel(lng父号) = 1
                End If
                
                If Val(.RowData(lngRow - 1)) <> 1 Then
                    .IsSubtotal(lngRow - 1) = False
                    .RowOutlineLevel(lngRow - 1) = 2
                End If
                .IsSubtotal(lngRow) = True
                .RowOutlineLevel(lngRow) = 2
            End If
            rsTemp.MoveNext
        Loop
    End With
    LoadWholeSubItems = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckISGetSubItem(lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断该行是否应该取从属项目(仅该行收费项目有从属项目及尚未取才取。)
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-31 10:41:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strSQL As String
    On Error GoTo ErrHandle
    strSQL = "Select count(从项ID) as 从属个数 From 收费从属项目 Where 主项ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsWholeSet.Cell(flexcpData, lngRow, vsWholeSet.ColIndex("收费项目"))))
    If rsTemp.EOF Then
        CheckISGetSubItem = False: Exit Function
    End If
    If Val(NVL(rsTemp!从属个数)) = 0 Then
        CheckISGetSubItem = False: Exit Function
    End If
    With vsWholeSet
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("从属父号"))) = Val(.Cell(flexcpData, lngRow, .ColIndex("序号"))) Then
                CheckISGetSubItem = False: Exit Function
            End If
        Next
    End With
    CheckISGetSubItem = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ShowSelectDept(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的执行部门
    '入参:strInput-输入的检查串
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-31 16:40:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, strKey As String
    Dim str类别 As String, lng执行科室 As Long, strWhere As String, lng项目id As Long
    Dim sngX As Single, sngY As Single, lngH As Long, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With vsWholeSet
        If .Col <> .ColIndex("缺省执行科室") Then Exit Function
        If Val(.Cell(flexcpData, .Row, .ColIndex("收费项目"))) = 0 Then Exit Function
        str类别 = Trim(.TextMatrix(.Row, .ColIndex("类别")))
        If InStr(",4,5,6,7,", "," & str类别 & ",") > 0 Then Exit Function
        '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
        lng执行科室 = Val(.TextMatrix(.Row, .ColIndex("执行科室")))
        If lng执行科室 <> 0 And lng执行科室 <> 4 Then Exit Function
        lng项目id = Val(.Cell(flexcpData, .Row, .ColIndex("收费项目")))

        strKey = gstrLike & strInput & "%"
        strWhere = ""
        If strInput <> "" Then
            If IsNumeric(strInput) Then
                strWhere = " And A.编码 Like [3]"
            ElseIf zlStr.IsCharAlpha(strInput) Then
                strWhere = " And A.简码 Like upper([3])"
            Else
                strWhere = " And (A.编码 Like [3] or A.简码 Like upper([3]) or A.名称 like [3] )"
            End If
        End If
            
        '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
        If lng执行科室 = 0 Then
            strSQL = _
            "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
            " From 部门表 A,部门性质说明 B " & _
            " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            "       And B.部门ID=A.ID and B.服务对象 IN(2,3) " & strWhere & _
            " Order by B.服务对象,A.编码"
        Else  '4
            strSQL = "" & _
            " Select Distinct A.ID,A.编码, A.名称" & _
            " From 收费执行科室 B,部门表 A" & _
            " Where B.收费细目ID=[1] And B.执行科室ID=A.id " & strWhere & _
            "       And (b.病人来源 is NULL Or b.病人来源=[2]) " & _
            " Order by A.编码" '
        End If
    End With
    Call CalcPosition(sngX, sngY, vsWholeSet)
     lngH = vsWholeSet.CellHeight
     sngY = sngY - lngH
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行部门选择", False, "", "", False, False, _
        True, sngX, sngY, lngH, blnCancel, False, True, lng项目id, 2, strKey)
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "未找到匹配的执行科室,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "未找到匹配的执行科室,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    With vsWholeSet
        .TextMatrix(.Row, .ColIndex("缺省执行科室")) = rsTemp!编码 & "-" & rsTemp!名称
        .Cell(flexcpData, .Row, .ColIndex("缺省执行科室")) = NVL(rsTemp!ID)
    End With
    ShowSelectDept = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsEdit执行科室() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许编辑执行科室
    '返回:允许返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-31 17:09:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, strKey As String
    Dim str类别 As String, lng执行科室 As Long
    With vsWholeSet
      If .Col <> .ColIndex("缺省执行科室") Then Exit Function
        If Val(.Cell(flexcpData, .Row, .ColIndex("收费项目"))) = 0 Then Exit Function
        str类别 = Trim(.TextMatrix(.Row, .ColIndex("类别")))
        If InStr(",4,5,6,7,", "," & str类别 & ",") > 0 Then Exit Function
        '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
        lng执行科室 = Val(.TextMatrix(.Row, .ColIndex("执行科室")))
        If lng执行科室 <> 0 And lng执行科室 <> 4 Then Exit Function
        IsEdit执行科室 = True
    End With
End Function
Private Sub FillBillComboBox(lngRow As Long, lngCol As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将相关执行科室加载到指定的combox中
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-31 17:05:23
    '说明:暂未用该过程,主要是科室多了不利于用户选择
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str类别 As String, lng执行科室 As Long
    
    On Error GoTo ErrHandle
    With vsWholeSet
        If lngCol <> .ColIndex("缺省执行科室") Then Exit Sub
        
        If Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目"))) <> 0 Then
              str类别 = Trim(.TextMatrix(lngRow, .ColIndex("类别")))
              If InStr(",4,5,6,7,", "," & str类别 & ",") > 0 Then
                    '药品,卫材
                      .ColComboList(.ColIndex("缺省执行科室")) = ""
              Else
                  '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
                  lng执行科室 = Val(.TextMatrix(lngRow, .ColIndex("执行科室")))
                  Select Case lng执行科室
                  Case 0
                         .ColComboList(.ColIndex("缺省执行科室")) = .BuildComboList(mrsDept, "名称", "ID", vbRed)
                  Case 4
                        strSQL = "Select Distinct b.ID,B.编码, B.名称" & _
                            " From 收费执行科室 A,部门表 B" & _
                            " Where A.收费细目ID=[1] And A.执行科室ID=b.id" & _
                            "       And (病人来源 is NULL Or 病人来源=[2]) " & _
                            " Order by B.编码" '
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Cell(flexcpData, lngRow, .ColIndex("收费项目"))), 2)
                         .ColComboList(.ColIndex("缺省执行科室")) = .BuildComboList(rsTmp, "名称", "ID", vbRed)
                  Case Else
                      .ColComboList(.ColIndex("缺省执行科室")) = ""
                  End Select
              End If
        Else
            .ColComboList(.ColIndex("缺省执行科室")) = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ItemExist(ByVal lng中药ID As Long, ByVal lngRow As Long) As Boolean
    '功能：判断中药配方输入表格中,指定的中药是否已经输入
    Dim i As Long, j As Long, lngTemp As Long
    With vsWholeSet
        For i = 1 To .Rows - 1
            If i <> lngRow Then
                If Val(.TextMatrix(i, .ColIndex("药名ID"))) = lng中药ID Then
                       ItemExist = True
                       Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'功能：判断输入金额是否在原价和现从限定的范围内
'参数：varL=原价,varR=现价,varI=输入金额
'返回：如果不在范围内,则为提示信息,否则为空串
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '如果数值符号相同,则用绝对值判断
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "输入的价格绝对值不在范围(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")内."
        End If
    Else
        '如果符号不相同,则用原始范围判断
        If varI < varL Or varI > varR Then
            CheckScope = "输入的价格值不在范围(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")内."
        End If
    End If
End Function
