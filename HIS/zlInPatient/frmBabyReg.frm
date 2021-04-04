VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBabyReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "新生儿登记"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmBabyReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9915
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&S)"
      Height          =   350
      Left            =   2520
      TabIndex        =   45
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelivery 
      Caption         =   "分娩信息(&F)"
      Height          =   350
      Left            =   5160
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame fraSplit 
      Height          =   75
      Left            =   0
      TabIndex        =   43
      Top             =   4880
      Width           =   10680
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8640
      TabIndex        =   16
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7440
      TabIndex        =   15
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印腕带(&P)"
      Height          =   350
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Height          =   350
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   1320
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.Frame fraMotherInfo 
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   9660
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "25岁x"
         Height          =   180
         Index           =   2
         Left            =   7920
         TabIndex        =   42
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "三病区x"
         Height          =   180
         Index           =   4
         Left            =   4440
         TabIndex        =   41
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "孕妇1x"
         Height          =   180
         Index           =   1
         Left            =   4440
         TabIndex        =   40
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "妇产科xx"
         Height          =   180
         Index           =   3
         Left            =   960
         TabIndex        =   39
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         Caption         =   "20101118xx"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   38
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科  室："
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lbl病区 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病  区："
         Height          =   180
         Left            =   3720
         TabIndex        =   34
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年  龄："
         Height          =   180
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓  名："
         Height          =   180
         Left            =   3720
         TabIndex        =   32
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl标识号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号："
         Height          =   180
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBaby 
      Height          =   1845
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   9660
      _cx             =   17039
      _cy             =   3254
      Appearance      =   3
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
      BackColorSel    =   16444122
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBabyReg.frx":058A
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
   Begin VB.Frame fraBabyInput 
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   9660
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   10
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   9
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   3
         Left            =   4560
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   0
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   8
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   7
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Index           =   6
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtBaby 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   5
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox txtBaby 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   1
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   1
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox txtBaby 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   4
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   7
         Top             =   360
         Width           =   810
      End
      Begin VB.TextBox txtBaby 
         BackColor       =   &H8000000E&
         Height          =   300
         Index           =   2
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "死亡时间"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   46
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblERRInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厘米"
         Height          =   180
         Index           =   9
         Left            =   8760
         TabIndex        =   37
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "克"
         Height          =   180
         Index           =   10
         Left            =   8760
         TabIndex        =   36
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注说明"
         Height          =   180
         Index           =   12
         Left            =   3720
         TabIndex        =   29
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生时间"
         Height          =   180
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "血  型"
         Height          =   180
         Index           =   7
         Left            =   7200
         TabIndex        =   27
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "胎儿状况"
         Height          =   180
         Index           =   4
         Left            =   3720
         TabIndex        =   26
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婴儿姓名"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "次"
         Height          =   180
         Index           =   8
         Left            =   1920
         TabIndex        =   23
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体  重"
         Height          =   180
         Index           =   6
         Left            =   7200
         TabIndex        =   22
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身  长"
         Height          =   180
         Index           =   5
         Left            =   7200
         TabIndex        =   21
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分娩方式"
         Height          =   180
         Index           =   3
         Left            =   3720
         TabIndex        =   20
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分娩次数"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婴儿性别"
         Height          =   180
         Index           =   1
         Left            =   3720
         TabIndex        =   18
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   0
      X2              =   8280
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "frmBabyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_BABY = 9
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mbln门诊 As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean
Private mstrPrivs As String
Private marrDelBaby() As Variant
Private mblnWristletPrint As Boolean    '是否打印病人腕带
Private mfrmParent As Object
Private mcolBaby As Collection     '便于根据索引值定位显示列 KEY(索引):VALUE（_列号）

Private Const M_CON_ColorUnEnabled = &H80000016
Private Const M_CON_ColorEnabled = &H8000000E

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Enum mCol
    col序号 = 0
    Col婴儿姓名 = 1
    Col婴儿性别 = 2
    Col分娩次数 = 3
    Col分娩方式 = 4
    Col胎儿状况 = 5
    Col身长 = 6
    Col体重 = 7
    COl血型 = 8
    Col出生时间 = 9
    Col死亡时间 = 10
    Col备注说明 = 11
End Enum

Private Enum M_E_SHOW
    IX_标识号 = 0
    IX_姓名 = 1
    IX_年龄 = 2
    IX_科室 = 3
    IX_病区 = 4
End Enum

Private Enum M_E_BABY
    B_姓名 = 0
    B_分娩次数 = 1
    B_出生时间 = 2
    B_备注说明 = 3
    B_身长 = 4
    B_体重 = 5
    
    B_性别 = 6
    B_分娩方式 = 7
    B_胎儿状况 = 8
    B_血型 = 9
    B_死亡时间 = 10
End Enum

Public Function ShowMe(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, strPrivs As String, frmParent As Object, Optional bln门诊 As Boolean) As Boolean
'参数：lng就诊ID=住院病人为主页ID，门诊病人为挂号ID。
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mbln门诊 = bln门诊
    mstrPrivs = strPrivs
    Set mfrmParent = frmParent
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cboBaby_Click(Index As Integer)
    If Me.Visible Then
        ChangeBabyInfo vsBaby.Row, mcolBaby("_" & Index), cboBaby(Index)
    End If
End Sub

Private Sub cboBaby_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ChangeBabyInfo(vsBaby.Row, mcolBaby("_" & Index), cboBaby(Index))
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
End Sub

Private Sub cmdAdd_Click()
    Call AddNewBabyRow
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH

    With vsBaby
        If .Rows = .FixedRows Or .Row <= .FixedRows - 1 Then Exit Sub
        If .RowData(.Row) <> 0 Then
            '对婴儿业务数据进行检查
            If mbln门诊 Then
                'by lesfeng 2009-12-29 大表拆分  病人费用记录 --〉门诊费用记录 这里只对门诊 "没有主页id 去掉And A.主页ID is Null"
                strSQL = _
                    " Select Distinct 1 as 标志,A.婴儿费 as 婴儿 From 门诊费用记录 A,病人挂号记录 B" & _
                    "   Where A.病人ID=[1]  And B.ID=[2] And A.婴儿费>=[3] And A.登记时间>=B.登记时间 and B.记录性质=1 and B.记录状态=1 And A.记录状态 =1 " & _
                    " Union ALL Select Distinct 2,A.婴儿 From 病人医嘱记录 A,病人挂号记录 B Where A.病人ID=[1] And A.挂号单=B.NO And B.ID=[2] And A.婴儿>=[3] And B.记录性质=1 and B.记录状态=1 And A.医嘱状态 <> 4" & _
                    " Union ALL Select Distinct 3,婴儿 From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3]" & _
                    " Union ALL Select Distinct 4,婴儿 From 病人护理记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3]" & _
                    " Union ALL Select Distinct 4,婴儿 From 病人护理文件 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3]"
            Else
                'by lesfeng 2009-12-29 大表拆分  病人费用记录 --〉住院费用记录 这里只对住院
                strSQL = _
                    " Select Distinct 1 as 标志,婴儿费 as 婴儿 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿费>=[3] And 记录状态 = 1" & _
                    " Union ALL Select Distinct 2,婴儿 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3] And 医嘱状态 <> 4" & _
                    " Union ALL Select Distinct 3,婴儿 From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3]" & _
                    " Union ALL Select Distinct 4,婴儿 From 病人护理记录 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3]" & _
                    " Union ALL Select Distinct 4,婴儿 From 病人护理文件 Where 病人ID=[1] And 主页ID=[2] And 婴儿>=[3]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID, Val(.RowData(.Row)))
            If Not rsTmp.EOF Then
                MsgBox "该病人的婴儿" & rsTmp!婴儿 & "已存在有效的" & Decode(rsTmp!标志, 1, "费用", 2, "医嘱", 3, "病历", 4, "护理") & "数据，当前行不能删除。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If MsgBox("确实要删除婴儿" & .TextMatrix(.Row, mCol.col序号) & _
            IIf(.TextMatrix(.Row, mCol.Col婴儿姓名) <> "", """" & .TextMatrix(.Row, mCol.Col婴儿姓名) & """", "") & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        For i = .Row + 1 To .Rows - 1
            .TextMatrix(i, mCol.col序号) = .TextMatrix(i, mCol.col序号) - 1
        Next
        i = .Row
        ReDim Preserve marrDelBaby(UBound(marrDelBaby) + 1)
        marrDelBaby(UBound(marrDelBaby)) = .RowData(.Row)
        SetBabyInfo .Row, 2   '清除卡片信息
        .RemoveItem .Row
        .Row = IIf(i <= .Rows - 1, i, .Rows - 1)
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdDelivery_Click()
    Dim objMedRecPage As zlMedRecPage.clsInOutMedRec
    
    Set objMedRecPage = New zlMedRecPage.clsInOutMedRec
    Call objMedRecPage.InitMedRec(gcnOracle, glngSys, glngModul)
    Call objMedRecPage.EditDelivery(Me, mlng病人ID, mlng就诊ID)
End Sub

Private Sub cmdOK_Click()
    Dim arrSQL As Variant, arrBaby As Variant, arrItem As Variant
    Dim blnTrans As Boolean
    Dim str出生时间 As String
    Dim strSQL As String, strNO As String, strErr As String
    Dim intAddCount As Integer
    Dim i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim blnLis As Boolean
    Dim strDieDate As String

    If Not CheckBaby() Then Exit Sub
    blnLis = Sys.IsSysSetUp(2500)
    arrSQL = Array()
    If blnLis Then arrBaby = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人新生儿记录_Delete(" & mlng病人ID & "," & mlng就诊ID & ")"
    With vsBaby
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, mCol.Col婴儿姓名) = "" And .TextMatrix(i, mCol.Col婴儿性别) = "" _
                And .TextMatrix(i, mCol.Col分娩方式) = "" And .TextMatrix(i, mCol.Col胎儿状况) = "" _
                And .TextMatrix(i, mCol.Col身长) = "" And .TextMatrix(i, mCol.Col体重) = "" And .TextMatrix(i, mCol.COl血型) = "" Then
                MsgBox "婴儿" & .TextMatrix(i, mCol.col序号) & "的信息录入不完整。", vbInformation, gstrSysName
                .Row = i: .ShowCell .Row, .Col: .SetFocus: Exit Sub
            End If
            If .TextMatrix(i, mCol.Col出生时间) <> "" And .TextMatrix(i, mCol.Col死亡时间) <> "" Then
                If Format(.TextMatrix(i, mCol.Col出生时间), "YYYY-MM-dd HH:mm") > Format(.TextMatrix(i, mCol.Col死亡时间), "YYYY-MM-dd HH:mm") Then
                    MsgBox "婴儿" & .TextMatrix(i, mCol.col序号) & "的【死亡时间】不能小于【出生时间】。", vbInformation, gstrSysName
                    .Row = i: .ShowCell .Row, .Col: .SetFocus: Exit Sub
                End If
            End If
            If .TextMatrix(i, mCol.Col死亡时间) = "" Then
                strDieDate = ""
            Else
                strDieDate = ",to_date('" & Mid(.TextMatrix(i, mCol.Col死亡时间), 1, 10) & " " & Val(Mid(.TextMatrix(i, mCol.Col死亡时间), 12, 2)) & _
                ":" & Val(Mid(.TextMatrix(i, mCol.Col死亡时间), 15, 2)) & "','yyyy-mm-dd hh24:mi:ss')"
            End If
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人新生儿记录_Insert(" & _
                mlng病人ID & "," & mlng就诊ID & "," & .TextMatrix(i, mCol.col序号) & "," & _
                "'" & .TextMatrix(i, mCol.Col婴儿姓名) & "','" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.Col婴儿性别)) & "'," & _
                ZVal(.TextMatrix(i, mCol.Col分娩次数)) & ",'" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.Col分娩方式)) & "'," & _
                "'" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.Col胎儿状况)) & "',to_date('" & _
                Mid(.TextMatrix(i, mCol.Col出生时间), 1, 10) & " " & Val(Mid(.TextMatrix(i, mCol.Col出生时间), 12, 2)) & _
                ":" & Val(Mid(.TextMatrix(i, mCol.Col出生时间), 15, 2)) & "','yyyy-mm-dd hh24:mi:ss')," & ZVal(.TextMatrix(i, mCol.Col身长)) & _
                "," & ZVal(.TextMatrix(i, mCol.Col体重)) & ",'" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.COl血型)) & "','" & _
                .TextMatrix(i, mCol.Col备注说明) & "'" & strDieDate & ")"
            If blnLis Then
                ReDim Preserve arrBaby(UBound(arrBaby) + 1)
                arrBaby(UBound(arrBaby)) = .TextMatrix(i, mCol.col序号) & ";" & .TextMatrix(i, mCol.Col婴儿姓名) & ";" & zlCommFun.GetNeedName(.TextMatrix(i, mCol.Col婴儿性别))
            End If
            If .RowData(i) = "" Then intAddCount = intAddCount + 1
        Next
    End With
    
    On Error GoTo errH
    
    If blnLis Then
        '在提交事务之前初始化病人公共部件对象
        If CreatePublicPatient() Then
            If Not gobjPublicPatient.InitLis(True) Then Exit Sub
        Else
            Exit Sub
        End If
        
        If mbln门诊 Then
            strSQL = "Select a.NO  From 病人挂号记录 A Where a.Id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng就诊ID)
            If Not rsTmp.EOF Then strNO = rsTmp!NO & ""
        End If
        
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        For j = LBound(arrBaby) To UBound(arrBaby)
            arrItem = Split(arrBaby(j), ";")
            rsTmp.Filter = "名称='" & arrItem(2) & "'"
            If Not rsTmp.EOF Then arrItem(2) = rsTmp!编码
            arrBaby(j) = arrItem(0) & ";" & arrItem(1) & ";" & arrItem(2)    '每个婴儿对应的序号、姓名、性别
        Next
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    '存在婴儿医嘱需要同步修改LIS中婴儿的数据
    If blnLis Then
        For i = LBound(arrBaby) To UBound(arrBaby)
             arrItem = Split(CStr(arrBaby(i)), ";")
             If Not gobjPublicPatient.ModifyBabyInfo(mlng病人ID, IIf(mbln门诊, 0, mlng就诊ID), IIf(mbln门诊, strNO, ""), CLng(arrItem(0)), arrItem(1), arrItem(2), strErr) Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 MsgBox "LIS 系统中新生儿信息更新失败！" & vbCrLf & IIf(strErr <> "", "错误原因:" & strErr, ""), vbOKOnly + vbInformation, Me.Caption
                 Exit Sub
             End If
         Next
     End If
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error Resume Next
    
    If mbln门诊 = False Then
        For i = 0 To UBound(marrDelBaby)
            If mclsMipModule.IsConnect = True Then
                mclsXML.ClearXmlText '清除缓存中的XML
                'patient_id      病人id  1   N
                mclsXML.appendData "patient_id", mlng病人ID, xsNumber
                'page_id     主页id  1   N
                mclsXML.appendData "page_id", mlng就诊ID, xsNumber
                'baby_serial     序号    1   N
                mclsXML.appendData "baby_serial", CInt(marrDelBaby(i)), xsNumber
                mclsMipModule.CommitMessage "ZLHIS_PATIENT_013", mclsXML.XmlText
            End If
        Next i
        '新生儿登记或删除发送消息
        For i = vsBaby.FixedRows To vsBaby.Rows - 1
            If mclsMipModule.IsConnect = True And Val(vsBaby.TextMatrix(i, mCol.col序号)) <> Val(vsBaby.RowData(i)) Then
                '删除婴儿会导致后面婴儿的序号发生变化
                If Val(vsBaby.TextMatrix(i, mCol.col序号)) <> Val(vsBaby.RowData(i)) And Val(vsBaby.RowData(i)) <> 0 Then
                    mclsXML.ClearXmlText '清除缓存中的XML
                    'patient_id      病人id  1   N
                    mclsXML.appendData "patient_id", mlng病人ID, xsNumber
                    'page_id     主页id  1   N
                    mclsXML.appendData "page_id", mlng就诊ID, xsNumber
                    'baby_serial     序号    1   N
                    mclsXML.appendData "baby_serial", Val(vsBaby.RowData(i)), xsNumber
                    mclsMipModule.CommitMessage "ZLHIS_PATIENT_013", mclsXML.XmlText
                End If
                
                mclsXML.ClearXmlText '清除缓存中的XML
                'in_patient 1
                mclsXML.AppendNode "in_patient"
                'patient_id      病人id  1   N
                mclsXML.appendData "patient_id", mlng病人ID, xsNumber
                'page_id     主页id  1   N
                mclsXML.appendData "page_id", mlng就诊ID, xsNumber
                'patient_name        姓名    1   S
                mclsXML.appendData "patient_name", lblOut(IX_姓名).Caption, xsString
                'patient_sex     性别    0..1    S
                mclsXML.appendData "patient_sex", lblOut(IX_年龄).Tag, xsString
                'in_number       住院号  1   S
                mclsXML.appendData "in_number", lblOut(IX_标识号).Caption, xsString
                mclsXML.AppendNode "in_patient", True
                'patient_baby 1
                mclsXML.AppendNode "patient_baby"
                'baby_serial     序号    1   N
                mclsXML.appendData "baby_serial", Val(vsBaby.TextMatrix(i, mCol.col序号)), xsNumber
                'baby_name       姓名    1   S
                mclsXML.appendData "baby_name", vsBaby.TextMatrix(i, mCol.Col婴儿姓名), xsString
                'baby_sex        性别    0..1    S
                mclsXML.appendData "baby_sex", zlCommFun.GetNeedName(vsBaby.TextMatrix(i, mCol.Col婴儿性别)), xsString
                'baby_birth      出生日期    0..1    S
                str出生时间 = Format(Mid(vsBaby.TextMatrix(i, mCol.Col出生时间), 1, 10) & " " & Val(Mid(vsBaby.TextMatrix(i, mCol.Col出生时间), 12, 2)) & _
                    ":" & Val(Mid(vsBaby.TextMatrix(i, mCol.Col出生时间), 15, 2)), "YYYY-MM-DD HH:mm:ss")
                mclsXML.appendData "baby_birth", str出生时间, xsString
                mclsXML.AppendNode "patient_baby", True
                mclsMipModule.CommitMessage "ZLHIS_PATIENT_011", mclsXML.XmlText
            End If
        Next i
    End If
    If Err <> 0 Then Err.Clear
    
    On Error GoTo errH
    
    '打印病人腕带
    If InStr(mstrPrivs, "婴儿腕带打印") Then
        mblnWristletPrint = True
        If gbytBabyWristletPrint = 0 Then
            mblnWristletPrint = False
        Else
            If gbytBabyWristletPrint = 2 And intAddCount > 0 Then
                If MsgBox("是否打印病人腕带？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mblnWristletPrint = False
                End If
            End If
        End If
        
        If mblnWristletPrint Then

            With vsBaby
                For i = .FixedRows To .Rows - 1
                    If (.RowData(i) = "") Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me, "病人ID=" & mlng病人ID, "主页ID=" & mlng就诊ID, "序号=" & .TextMatrix(i, mCol.col序号), 2)
                    End If
                Next
            End With
        End If
    End If
    
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdPrint_Click()
    With vsBaby
        If .RowData(.Row) = "" Then MsgBox "新增新生儿需要在确定时才能打印腕带！", vbInformation + vbOKOnly, gstrSysName: Exit Sub
        
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me, "病人ID=" & mlng病人ID, "主页ID=" & mlng就诊ID, "序号=" & .TextMatrix(.Row, mCol.col序号), 2)
    End With
End Sub

Private Sub cmdPrintSet_Click()
'功能:腕带打印设置
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me)
End Sub

Private Sub Form_Activate()
    '触发行变动,使卡片值与当前行保持一致
    If vsBaby.Rows > 1 Then
        vsBaby.Row = 0
        vsBaby.Row = vsBaby.Rows - 1
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    mblnChange = False
    mblnOK = False
    marrDelBaby = Array()
    
    On Error GoTo errH
    
    '产妇信息
    If mbln门诊 Then
        lbl标识号.Caption = "门诊号"
        lbl科室.Caption = "科室"
        lbl病区.Visible = False: lblOut(IX_病区).Visible = False
        
        strSQL = "Select B.门诊号 as 标识号,B.姓名,B.性别,B.年龄,C.名称 as 科室,NULL as 病区" & _
            " From 病人挂号记录 B,部门表 C" & _
            " Where B.执行部门ID=C.ID And B.ID=[1] And B.记录性质=1 and B.记录状态=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng就诊ID)
    Else
        strSQL = "Select B.住院号 as 标识号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,B.年龄,C.名称 as 科室,D.名称 as 病区" & _
            " From 病人信息 A,病案主页 B,部门表 C,部门表 D" & _
            " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID And B.当前病区ID=D.ID" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    End If
    
    lblOut(IX_标识号).Caption = Nvl(rsTmp!标识号)
    lblOut(IX_姓名).Caption = Nvl(rsTmp!姓名)
    lblOut(IX_年龄).Caption = Nvl(rsTmp!年龄)
    lblOut(IX_年龄).Tag = Nvl(rsTmp!性别)
    lblOut(IX_科室).Caption = Nvl(rsTmp!科室)
    lblOut(IX_病区) = Nvl(rsTmp!病区)
    
    '婴儿信息
    strSQL = "Select 序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,身长,体重,血型,出生时间,死亡时间,备注说明" & _
        " From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    With vsBaby
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, mCol.col序号) = rsTmp!序号
                .TextMatrix(i, mCol.Col婴儿姓名) = Nvl(rsTmp!婴儿姓名)
                .TextMatrix(i, mCol.Col婴儿性别) = Nvl(rsTmp!婴儿性别)
                .TextMatrix(i, mCol.Col分娩次数) = Nvl(rsTmp!分娩次数)
                .TextMatrix(i, mCol.Col分娩方式) = Nvl(rsTmp!分娩方式)
                .TextMatrix(i, mCol.Col胎儿状况) = Nvl(rsTmp!胎儿状况)
                .TextMatrix(i, mCol.Col身长) = gclsBase.FormatEx(rsTmp!身长, 2)  '厘米
                .TextMatrix(i, mCol.Col体重) = gclsBase.FormatEx(rsTmp!体重, 2)  '克
                .TextMatrix(i, mCol.COl血型) = Nvl(rsTmp!血型)
                .TextMatrix(i, mCol.Col出生时间) = Format(rsTmp!出生时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, mCol.Col死亡时间) = Format(Nvl(rsTmp!死亡时间), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, mCol.Col备注说明) = Nvl(rsTmp!备注说明)
                .RowData(i) = Val(rsTmp!序号) '表明是已有数据
                
                '隐藏数据
                .Cell(flexcpData, i, mCol.Col婴儿姓名) = Nvl(rsTmp!婴儿姓名)
                .Cell(flexcpData, i, mCol.Col婴儿性别) = Nvl(rsTmp!婴儿性别)
                .Cell(flexcpData, i, mCol.Col分娩次数) = Nvl(rsTmp!分娩次数)
                .Cell(flexcpData, i, mCol.Col分娩方式) = Nvl(rsTmp!分娩方式)
                .Cell(flexcpData, i, mCol.Col胎儿状况) = Nvl(rsTmp!胎儿状况)
                .Cell(flexcpData, i, mCol.Col身长) = gclsBase.FormatEx(rsTmp!身长, 2)
                .Cell(flexcpData, i, mCol.Col体重) = gclsBase.FormatEx(rsTmp!体重, 2)
                .Cell(flexcpData, i, mCol.COl血型) = Nvl(rsTmp!血型)
                .Cell(flexcpData, i, mCol.Col出生时间) = Format(rsTmp!出生时间, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, mCol.Col死亡时间) = Format(Nvl(rsTmp!死亡时间), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, mCol.Col备注说明) = Nvl(rsTmp!备注说明)
                rsTmp.MoveNext
            Next
        Else
            .Rows = 1
        End If

        Call SetCardEnable(.Rows > 1)
    End With
    
    
    '性别选择
    Call ReadDict("性别", cboBaby(B_性别))
    '分娩方式选择
    Call ReadDict("分娩方式", cboBaby(B_分娩方式))
    '胎儿状况选择
    Call ReadDict("胎儿状况", cboBaby(B_胎儿状况))
    '血型加载
    Call ReadDict("血型", cboBaby(B_血型))
    
    '通过卡片索引找到列号
    Set mcolBaby = New Collection
    
    With mcolBaby
        .Add mCol.Col婴儿姓名, "_" & B_姓名
        .Add mCol.Col分娩次数, "_" & B_分娩次数
        .Add mCol.Col出生时间, "_" & B_出生时间
        .Add mCol.Col备注说明, "_" & B_备注说明
        .Add mCol.Col身长, "_" & B_身长
        .Add mCol.Col体重, "_" & B_体重
        .Add mCol.Col分娩方式, "_" & B_分娩方式
        .Add mCol.Col胎儿状况, "_" & B_胎儿状况
        .Add mCol.COl血型, "_" & B_血型
        .Add mCol.Col婴儿性别, "_" & B_性别
        .Add mCol.Col死亡时间, "_" & B_死亡时间
        '--通过列号找到卡片对象
        .Add B_姓名, "C" & mCol.Col婴儿姓名
        .Add B_分娩次数, "C" & mCol.Col分娩次数
        .Add B_出生时间, "C" & mCol.Col出生时间
        .Add B_备注说明, "C" & mCol.Col备注说明
        .Add B_身长, "C" & mCol.Col身长
        .Add B_体重, "C" & mCol.Col体重
        .Add B_分娩方式, "C" & mCol.Col分娩方式
        .Add B_胎儿状况, "C" & mCol.Col胎儿状况
        .Add B_血型, "C" & mCol.COl血型
        .Add B_性别, "C" & mCol.Col婴儿性别
        .Add B_死亡时间, "C" & mCol.Col死亡时间
    End With
    
    cmdPrint.Visible = InStr(mstrPrivs, "婴儿腕带打印") > 0
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("数据已经被修改，确实要不保存退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    Set mcolBaby = Nothing
    
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub txtBaby_Change(Index As Integer)
    vsBaby.TextMatrix(vsBaby.Row, mcolBaby("_" & Index)) = txtBaby(Index).Text
    If vsBaby.Cell(flexcpData, vsBaby.Row, mcolBaby("_" & Index)) <> vsBaby.TextMatrix(vsBaby.Row, mcolBaby("_" & Index)) Then mblnChange = True
End Sub

Private Sub txtBaby_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtBaby(Index)
End Sub

Private Sub txtBaby_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = Asc("'") Then
        KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        '回车13
        KeyAscii = 0
        Select Case Index
        
        Case B_姓名
            If Trim(txtBaby(Index).Text) = "" Then
                txtBaby(Index).Text = Trim(lblOut(IX_姓名).Caption & "之婴" & vsBaby.TextMatrix(vsBaby.Row, mCol.col序号))
            End If
        Case B_分娩次数
            If Trim(txtBaby(Index).Text) = "" Then
                txtBaby(Index).Text = 1
            End If
        Case B_出生时间
            If Trim(txtBaby(Index).Text) = "" Then
                txtBaby(Index).Text = Format(zlDatabase.Currentdate, "YYYY-MM-dd HH:mm")
            ElseIf Trim(txtBaby(Index).Text) <> "" Then
                txtBaby(Index).Text = GetFullDate(txtBaby(Index).Text)
            End If
        Case B_死亡时间
            If Trim(txtBaby(Index).Text) <> "" Then
                txtBaby(Index).Text = GetFullDate(txtBaby(Index).Text)
            End If
        Case B_身长, B_体重
            txtBaby(Index).Text = gclsBase.FormatEx(txtBaby(Index).Text, 2) '最多保留两位
        End Select
        Call ChangeBabyInfo(vsBaby.Row, mcolBaby("_" & Index), txtBaby(Index))
        If Index = B_备注说明 Then
            If vsBaby.Row = vsBaby.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                vsBaby.Row = vsBaby.Row + 1
                txtBaby(B_姓名).SetFocus: Exit Sub
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        Exit Sub
    End If
    
    If KeyAscii = vbKeyBack Then
        vsBaby.TextMatrix(vsBaby.Row, mcolBaby("_" & Index)) = ""
    End If
    
    Select Case Index
    Case B_出生时间, B_死亡时间
        If InStr("/-0123456789:" & Chr(32) & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Case B_分娩次数
        If InStr("0123456789" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    Case B_身长, B_体重
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txtBaby_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTmp As String
    
    strTmp = txtBaby(Index).Text
    If Index = B_姓名 Then
        If LenB(StrConv(strTmp, vbFromUnicode)) > 16 Then
            zlCommFun.ShowTipInfo txtBaby(Index).hWnd, strTmp
        Else
            zlCommFun.ShowTipInfo txtBaby(Index).hWnd, ""
        End If
        
    End If

    If Index = B_出生时间 Or Index = B_姓名 Then
        If strTmp = "" Then
            txtBaby(Index).ToolTipText = "回车设置缺省值"
        Else
            txtBaby(Index).ToolTipText = ""
        End If
    End If
End Sub

Private Sub txtBaby_Validate(Index As Integer, Cancel As Boolean)
    Dim re As New RegExp
    Dim strMsg As String
    
    Select Case Index
    Case B_姓名
        If zlCommFun.ActualLen(txtBaby(Index).Text) > 100 Then
            strMsg = "【婴儿姓名】最大只允许输入50个汉字或100个字符！"
        End If
    Case B_出生时间
        re.Pattern = "^([1-9][0-9]{3})-((01|03|05|07|08|10|12)-(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)-(0[1-9]|[1-2][0-9]|30)|02-(0[1-9]|[1-2][0-9]))\s([0-1][0-9]|2[0-3]):([0-5][0-9])$"
        If Not re.Test(Trim(txtBaby(Index).Text)) Then
            strMsg = "【出生时间】不是有效的日期格式[YYYY-MM-dd hh:mm]！"
            txtBaby(Index).SetFocus
        Else
            If CDate(Format(Trim(txtBaby(Index).Text), "YYYY-MM-dd HH:mm")) > CDate(Format(zlDatabase.Currentdate, "YYYY-MM-dd HH:mm")) Then
                strMsg = "【出生时间】大于当前系统时间！"
            End If
        End If
    Case B_死亡时间
        If Trim(txtBaby(Index).Text) <> "" Then
            re.Pattern = "^([1-9][0-9]{3})-((01|03|05|07|08|10|12)-(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)-(0[1-9]|[1-2][0-9]|30)|02-(0[1-9]|[1-2][0-9]))\s([0-1][0-9]|2[0-3]):([0-5][0-9])$"
            If Not re.Test(Trim(txtBaby(Index).Text)) Then
                strMsg = "【死亡时间】不是有效的日期格式[YYYY-MM-dd hh:mm]！"
                txtBaby(Index).SetFocus
            End If
        End If
    Case B_身长, B_体重
        If txtBaby(Index).Text = "" Then
            Exit Sub  '兼容以前没有录入身长\体重的情况
        ElseIf Not IsNumeric(txtBaby(Index).Text) Then
            strMsg = IIf(Index = B_身长, "【身长】", "【体重】") & "不是有效的数字类型！"
        ElseIf Len(txtBaby(Index).Text) > 10 Then
            strMsg = IIf(Index = B_身长, "【身长】", "【体重】") & "最大只允许录入10个字符！"
        End If
    Case B_备注说明
        If zlCommFun.ActualLen(txtBaby(Index).Text) > 100 Then
            strMsg = "【备注说明】最大只允许输入50个汉字或100个字符！"
        End If
    End Select
    
    If strMsg <> "" Then
        ShowErrInfo "提示:" & strMsg
        zlControl.TxtSelAll txtBaby(Index)
        Cancel = True: Exit Sub
    Else
        ShowErrInfo strMsg
    End If
End Sub

Private Sub vsBaby_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngTmp As Long
    
    With vsBaby
        If .Rows > 1 Then
            If NewRow <> OldRow And NewRow > 0 Then
                Call SetBabyInfo(NewRow)
            End If
        Else
            Call SetBabyInfo(0, 2)
        End If
        Call SetCardEnable(.Rows > 1)
    End With
End Sub

Private Sub vsBaby_Click()
    With vsBaby
        If .Rows > 1 Then
            
        End If
    End With
End Sub

Private Sub vsBaby_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call cmdDel_Click
    End If
End Sub

Private Function AddNewBabyRow() As Boolean
'功能：新增一数据行，并设置缺省值
'返回：如果不允许新增行，则提示并返回False
    Dim lngRow As Long
    Dim strMsg As String
    Dim i As Long
    
    
    With vsBaby
        If .Rows - 1 >= MAX_BABY Then
            MsgBox "病人的婴儿数太多，不能再随意增加。", vbInformation, gstrSysName
            Exit Function
        End If
        For i = 1 To .Rows - 1
            If .TextMatrix(i, mCol.Col婴儿姓名) = "" Then
                .Row = i
                txtBaby(B_姓名).SetFocus
                Exit Function
            End If
        Next
        
        .AddItem "", .Rows
        .Row = .Rows - 1: Call SetBabyInfo(.Row, 1)
        .TextMatrix(.Row, mCol.col序号) = .Rows - 1
        .ShowCell .Row, mCol.col序号
        txtBaby(B_姓名).SetFocus
        mblnChange = True
    End With
    
    AddNewBabyRow = True
End Function

Private Function ReadDict(strDict As String, cbo As ComboBox) As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbo.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo.ListIndex = cbo.NewIndex
                cbo.ItemData(cbo.NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
    End If
    ReadDict = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ChangeBabyInfo(ByVal lngRow As Long, ByVal lngCol As Long, ByVal objControl As Object)
    Dim strTmp As String

    If lngRow < vsBaby.FixedRows Then Exit Sub
    If TypeName(objControl) = "ComboBox" Then
        strTmp = zlCommFun.GetNeedName(Trim(objControl.Text))
    Else
        strTmp = Trim(objControl.Text)
    End If
    vsBaby.TextMatrix(lngRow, lngCol) = strTmp
End Sub

Private Sub SetBabyInfo(ByVal lngRow As Long, Optional ByVal bytFunc As Byte = 0)
'功能:将表格内容显示到卡片
'参数:bytFunc =0,选中行内容显示到卡片；=1新增设置缺省值=2清除卡片信息
    With vsBaby
        If lngRow = 0 Then Exit Sub
        If bytFunc = 0 Then
            On Error Resume Next
            '兼容以前没有血型的赋值会报错cboBaby(B_血型).Text赋值时,找不到值
            txtBaby(B_姓名).Text = .TextMatrix(lngRow, mCol.Col婴儿姓名)   '缺省姓名
            txtBaby(B_身长).Text = .TextMatrix(lngRow, mCol.Col身长)
            txtBaby(B_体重).Text = .TextMatrix(lngRow, mCol.Col体重)
            txtBaby(B_分娩次数).Text = .TextMatrix(lngRow, mCol.Col分娩次数)
            txtBaby(B_出生时间).Text = .TextMatrix(lngRow, mCol.Col出生时间)
            txtBaby(B_死亡时间).Text = .TextMatrix(lngRow, mCol.Col死亡时间)
            cboBaby(B_胎儿状况).Text = cbo.Locate(cboBaby(B_胎儿状况), .TextMatrix(lngRow, mCol.Col胎儿状况), False)
            cboBaby(B_血型).Text = cbo.Locate(cboBaby(B_血型), .TextMatrix(lngRow, mCol.COl血型), False)
            cboBaby(B_分娩方式).Text = cbo.Locate(cboBaby(B_分娩方式), .TextMatrix(lngRow, mCol.Col分娩方式), False)
            cboBaby(B_性别).Text = cbo.Locate(cboBaby(B_性别), .TextMatrix(lngRow, mCol.Col婴儿性别), False)
            
            txtBaby(B_备注说明).Text = .TextMatrix(lngRow, mCol.Col备注说明)
            
            Err.Clear: On Error GoTo 0
        ElseIf bytFunc = 1 Then
            .TextMatrix(lngRow, mCol.Col婴儿姓名) = txtBaby(B_姓名).Text
            txtBaby(B_分娩次数).Text = 1
            .TextMatrix(lngRow, mCol.Col分娩次数) = txtBaby(B_分娩次数).Text
            txtBaby(B_出生时间).Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
            .TextMatrix(lngRow, mCol.Col出生时间) = txtBaby(B_出生时间).Text
            .TextMatrix(lngRow, mCol.Col死亡时间) = txtBaby(B_死亡时间).Text
            .TextMatrix(lngRow, mCol.Col身长) = txtBaby(B_身长).Text
            .TextMatrix(lngRow, mCol.Col体重) = txtBaby(B_体重).Text
             
            .TextMatrix(lngRow, mCol.Col婴儿性别) = zlCommFun.GetNeedName(cboBaby(B_性别).Text)
            .TextMatrix(lngRow, mCol.Col分娩方式) = zlCommFun.GetNeedName(cboBaby(B_分娩方式).Text)
            .TextMatrix(lngRow, mCol.Col胎儿状况) = zlCommFun.GetNeedName(cboBaby(B_胎儿状况).Text)
            .TextMatrix(lngRow, mCol.COl血型) = zlCommFun.GetNeedName(cboBaby(B_血型).Text)
            
        ElseIf bytFunc = 2 Then
        '清除卡片信息
            txtBaby(B_姓名).Text = ""
            txtBaby(B_分娩次数).Text = ""
            txtBaby(B_身长).Text = ""
            txtBaby(B_体重).Text = ""
            txtBaby(B_备注说明).Text = ""
            txtBaby(B_出生时间).Text = ""
            txtBaby(B_死亡时间).Text = ""
        End If
    End With
End Sub
Private Function CheckBaby() As Boolean
'功能:新增前检查
    Dim i As Long, k As Long
    Dim strErr As String
    Dim j As Long
    Dim strName As String
    
    '检查录入信息不能为空
    With vsBaby
        strErr = ""
        For i = .FixedRows To .Rows - 1
            For j = mCol.Col婴儿姓名 To mCol.COl血型
                '兼容以前没有身长体重时,不要求必须录入
                If j = mCol.Col身长 Or j = mCol.Col体重 Then Exit For
                
                If .TextMatrix(i, j) = "" Then
                    strErr = "序号【" & .TextMatrix(i, mCol.col序号) & "】的" & .TextMatrix(0, j) & "为空！"
                    Exit For
                End If
            Next
            If strErr <> "" Then Exit For
            
            For k = 1 To .Rows - 1
                If .TextMatrix(i, mCol.Col婴儿姓名) = .TextMatrix(k, mCol.Col婴儿姓名) And k <> i Then
                    strErr = "该婴儿：【" & .TextMatrix(i, mCol.Col婴儿姓名) & "】重复添加。": j = mCol.Col婴儿姓名
                    Exit For
                End If
                
                If .TextMatrix(i, mCol.Col分娩次数) <> .TextMatrix(k, mCol.Col分娩次数) Then
                    strErr = "婴儿【" & .TextMatrix(i, mCol.Col婴儿姓名) & "】分娩次数与婴儿【" & .TextMatrix(k, mCol.Col婴儿姓名) & "】分娩次数不一致！": j = mCol.Col分娩次数
                    Exit For
                End If
            Next
            If k <= .Rows - 1 Then Exit For
            
        Next
        
        If i <= .Rows - 1 Then
            If strErr <> "" Then
                MsgBox strErr, vbInformation + vbOKOnly, Me.Caption
                 .Row = i: .ShowCell .Row, j
                If j = mCol.Col出生时间 Or j = mCol.Col婴儿姓名 Or j = mCol.Col分娩次数 Or j = mCol.Col身长 Or j = mCol.Col体重 Or j = mCol.Col备注说明 Then
                    txtBaby(mcolBaby("C" & j)).SetFocus
                Else
                    cboBaby(mcolBaby("C" & j)).SetFocus
                End If
             
                Exit Function
            End If
        End If
    End With
    
    CheckBaby = True
End Function

Private Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
'参数：blnTime=是否处理时间部份
    Dim Curdate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    Curdate = zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(Curdate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(Curdate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(Curdate, "yyyy-MM") & "-" & strTmp & " " & Format(Curdate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(Curdate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(Curdate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(Curdate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(Curdate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(Curdate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function

Private Sub SetCardEnable(ByVal blnEnable As Boolean)

    If fraBabyInput.Enabled = blnEnable Then Exit Sub
    cmdDel.Enabled = blnEnable
    cmdPrint.Enabled = blnEnable
    fraBabyInput.Enabled = blnEnable

    If blnEnable Then
        txtBaby(B_姓名).BackColor = M_CON_ColorEnabled
        txtBaby(B_分娩次数).BackColor = M_CON_ColorEnabled
        txtBaby(B_出生时间).BackColor = M_CON_ColorEnabled
        txtBaby(B_死亡时间).BackColor = M_CON_ColorEnabled
        txtBaby(B_身长).BackColor = M_CON_ColorEnabled
        txtBaby(B_体重).BackColor = M_CON_ColorEnabled
        txtBaby(B_备注说明).BackColor = M_CON_ColorEnabled
        
        cboBaby(B_性别).BackColor = M_CON_ColorEnabled
        cboBaby(B_血型).BackColor = M_CON_ColorEnabled
        cboBaby(B_分娩方式).BackColor = M_CON_ColorEnabled
        cboBaby(B_胎儿状况).BackColor = M_CON_ColorEnabled
        
    Else
        txtBaby(B_姓名).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_分娩次数).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_出生时间).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_死亡时间).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_身长).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_体重).BackColor = M_CON_ColorUnEnabled
        txtBaby(B_备注说明).BackColor = M_CON_ColorUnEnabled
        
        cboBaby(B_性别).BackColor = M_CON_ColorUnEnabled
        cboBaby(B_血型).BackColor = M_CON_ColorUnEnabled
        cboBaby(B_分娩方式).BackColor = M_CON_ColorUnEnabled
        cboBaby(B_胎儿状况).BackColor = M_CON_ColorUnEnabled
    End If
End Sub

Private Sub ShowErrInfo(ByVal strMsg As String)
    If strMsg = "" Then
        lblERRInfo.Visible = False
    Else
        lblERRInfo.Visible = True
        lblERRInfo.Caption = strMsg
    End If
End Sub
