VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiPrice 
   Caption         =   "病人划价单查找"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "frmPatiPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   10065
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboDiagnose 
      Height          =   300
      ItemData        =   "frmPatiPrice.frx":058A
      Left            =   960
      List            =   "frmPatiPrice.frx":0591
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1530
      Width           =   3500
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   5955
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiPrice.frx":059F
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraOk 
      Height          =   615
      Left            =   60
      TabIndex        =   23
      Top             =   5370
      Width           =   9855
      Begin VB.CommandButton cmdAllCls 
         Caption         =   "全清(&R)"
         Height          =   350
         Left            =   1140
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   195
         Width           =   945
      End
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8640
         TabIndex        =   25
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   7500
         TabIndex        =   24
         ToolTipText     =   "热键：F2"
         Top             =   165
         Width           =   1100
      End
   End
   Begin VB.Frame fraDays 
      Caption         =   "选择划价单"
      Height          =   1455
      Left            =   7965
      TabIndex        =   28
      Top             =   60
      Width           =   1830
      Begin VB.CheckBox chk缺省 
         Caption         =   "所有划价单"
         Height          =   360
         Index           =   2
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   990
         Width           =   1605
      End
      Begin VB.CheckBox chk缺省 
         Caption         =   "有效天数的划价单"
         Height          =   360
         Index           =   1
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   615
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.CheckBox chk缺省 
         Caption         =   "当日内划价单"
         Height          =   360
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   225
         Width           =   1605
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   1410
      Left            =   45
      TabIndex        =   20
      Top             =   3960
      Width           =   9825
      _cx             =   17330
      _cy             =   2487
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
   Begin VB.Frame fraHsc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      TabIndex        =   9
      Top             =   3825
      Width           =   9840
   End
   Begin VB.Frame fraPati 
      Caption         =   " 病人信息 "
      Height          =   1455
      Left            =   45
      TabIndex        =   8
      Top             =   60
      Width           =   7860
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   6525
         TabIndex        =   29
         Top             =   1020
         Width           =   1155
      End
      Begin VB.ComboBox cbo调整费别 
         Height          =   300
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox txt付款方式 
         Height          =   300
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txt门诊号 
         Height          =   300
         Left            =   750
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox txt年龄 
         Height          =   300
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txt性别 
         Height          =   300
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   1080
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   750
         MaxLength       =   100
         TabIndex        =   0
         Top             =   315
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3615
         TabIndex        =   7
         Top             =   1035
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483636
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   177012739
         CurrentDate     =   38073
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   750
         TabIndex        =   6
         Top             =   1035
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483636
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   177012739
         CurrentDate     =   38073
      End
      Begin VB.TextBox txt费别 
         Height          =   300
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label lbl调整费别 
         AutoSize        =   -1  'True
         Caption         =   "调整费别"
         Height          =   180
         Left            =   3750
         TabIndex        =   22
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   180
         Left            =   5745
         TabIndex        =   18
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原费别"
         Height          =   180
         Left            =   2070
         TabIndex        =   17
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4110
         TabIndex        =   15
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2250
         TabIndex        =   14
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   345
         TabIndex        =   12
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   180
         Left            =   3195
         TabIndex        =   11
         Top             =   1095
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
         Height          =   180
         Left            =   345
         TabIndex        =   10
         Top             =   1095
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   2070
      Left            =   0
      TabIndex        =   19
      Top             =   1830
      Width           =   9825
      _cx             =   17330
      _cy             =   3651
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Label lblDiagnose 
      BackStyle       =   0  'Transparent
      Caption         =   "按诊断过滤"
      Height          =   255
      Left            =   30
      TabIndex        =   34
      Top             =   1590
      Width           =   915
   End
End
Attribute VB_Name = "frmPatiPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrBills As String
Private mstrPrivs As String
Private mlngModule As String
Private mlng病人ID As Long
Private mrsList As ADODB.Recordset  '单据列表
Private mrsDetail As ADODB.Recordset
Private mbln不允许多单据 As Boolean
Private mlng挂号科室 As Long        '当通过挂号单输入时,传入病人当前挂号单的挂号科室
Private mblnCard As Boolean
Private mbln住院病人门诊收费 As Boolean '34182
Private mblnNotClick As Boolean
Private mblnPreCard As Boolean
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:是否缓存了回车键,可能存在在收费界面刷卡中本身包含了回车,因此需要判断

Public Function FindBill(frmParent As Object, _
    ByVal strPrivs As String, Optional ByVal lng病人ID As Long, _
    Optional ByVal bln不允许多单据 As Boolean, _
    Optional ByVal lng挂号科室 As Long, _
    Optional ByVal bln住院病人门诊收费 As Boolean = False, _
    Optional blnCard As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找划价单据
    '入参:lng病人ID=可以指定病人(该病人之前肯定已确定有划价单据)
    '        lng挂号科室,当通过挂号单输入时,传入病人当前挂号单的挂号科室
    '        bln住院病人门诊收费-住院病人按门诊进行收费:34182
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-19 14:37:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbln住院病人门诊收费 = bln住院病人门诊收费
    mstrPrivs = strPrivs
    mlng病人ID = lng病人ID: mblnPreCard = blnCard
    mbln不允许多单据 = bln不允许多单据
    mlng挂号科室 = lng挂号科室
    Me.Show 1, frmParent
    FindBill = mstrBills
End Function

Public Function GetPriceBillString(ByVal lng病人ID As Long, ByVal bln不允许多单据 As Boolean, ByVal lng挂号科室 As Long, _
    Optional ByVal bln住院病人门诊收费 As Boolean = False, Optional blnCard As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人指定划价时间范围内的划价单(不弹出窗体选择)
    '入参:bln不允许多单据,如果不允许多单据,只返回最后划价的一张单据号
    '        lng挂号科室,当通过挂号单输入时,传入病人当前挂号单的挂号科室
    '        bln住院病人门诊收费-住院病人按门诊进行收费:34182
    '       blnCard-是否读卡
    '出参:
    '返回:"G0001112,G0001113,G0001114..."
    '编制:刘兴洪
    '日期:2010-11-19 14:39:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, DatBegin As Date, DatEnd As Date
    Dim i As Long, strTmp As String
    mbln住院病人门诊收费 = bln住院病人门诊收费
    DatEnd = zlDatabase.Currentdate
    DatBegin = DatEnd - gintSeekDays
    mblnPreCard = blnCard
    Set rsTmp = GetPriceBills(lng病人ID, lng挂号科室, DatBegin, DatEnd)
    For i = 1 To rsTmp.RecordCount
        strTmp = strTmp & IIf(strTmp = "", "", ",") & rsTmp!单据号
        If bln不允许多单据 Then Exit For
        rsTmp.MoveNext
    Next
        
    If gblnCheckTest Then
        '只要存在药品皮试结果为阳性的都不允许收费
        If Not CheckTest(strTmp, DatBegin, DatEnd) Then strTmp = ""
    End If
    
    GetPriceBillString = strTmp

End Function
Private Sub Local费别(ByVal str费别 As String, Optional blnNotFindAdd As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费别，查找指定的费别
    '入参:blnNotFindAdd-没找到，直接增加
    '出参:
    '编制:刘兴洪
    '日期:2011-04-17 21:45:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo调整费别.ListCount - 1
        If str费别 = cbo调整费别.List(i) Then
            cbo调整费别.ListIndex = i: Exit Sub
        End If
    Next
    If blnNotFindAdd = False Then Exit Sub
    cbo调整费别.AddItem str费别
    cbo调整费别.ListIndex = cbo调整费别.NewIndex
End Sub

Private Sub cboDiagnose_Click()
    '74296,冉俊明,2014-7-4,按单据的诊断过滤,把单据中的诊断组成一个下拉列表供选择
    If mblnNotClick Then Exit Sub
    If mrsList Is Nothing Then Exit Sub
    mshList.Clear
    mshList.Rows = 2
    mshDetail.Clear
    mshDetail.Rows = 2
    stbThis.Panels(2).Text = ""
    mrsList.Filter = IIf(cboDiagnose.Text = "所有诊断", "", "诊断='" & cboDiagnose.Text & "'")
    Set mshList.DataSource = mrsList
    Call SetHeader
    Call SetDetail
    Call mshList_EnterCell
    stbThis.Panels(2).Text = GetBillNote
End Sub

Private Sub cbo调整费别_Click()
    Dim strSql As String, strComMand As String
    Dim i As Integer, strNos As String, blnChange As Boolean
    
    If mblnNotClick Then Exit Sub
    If cbo调整费别.ListIndex < 0 Then Exit Sub
    '79870:李南春,2015/4/10,调整部分单据的费别
    '因为可能存在部分单据与病人信息的费别不一致的情况，所以不再检查调整的费别是否与原费别相同
    If InStr(1, mstrPrivs, ";调整病人费别;") = 0 Then Exit Sub
    
    strComMand = zlCommFun.ShowMsgbox("注意", "你是否要将费别『" & Trim(txt费别.Text) & "』调整为『" & cbo调整费别.Text & "』吗?" & vbCrLf & "  调整费别将直接影响相关的收费划价单!" & vbCrLf & vbCrLf & _
    "『所有划价单』:将病人未收费的划价单全部按新的费别调整,医嘱新开的处方按新费别打折" & vbCrLf & vbCrLf & _
    "『选中划价单』:只调整被选中的划价单,不影响新开的处方和未选中的划价单" & vbCrLf & vbCrLf & _
    "『不调整』:不调整费别,还原设置" & vbCrLf, "所有划价单,选中划价单,不调整", Me, vbQuestion)
    Select Case strComMand
    Case "所有划价单"
    Case "选中划价单"
        With mshList
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("选择")) = "√" Then strNos = strNos & "," & .TextMatrix(i, .ColIndex("单据号"))
            Next
        End With
        If strNos = "" Then
            MsgBox "请先勾选需要调整的单据，再选择费别类型！", vbInformation, gstrSysName
            Exit Sub
        End If
        strNos = Mid(strNos, 2)
    Case Else
        mblnNotClick = True
        Local费别 txt费别.Text, True
        mblnNotClick = False
        Exit Sub
    End Select
    
    fraOk.Enabled = False: cmdAllSel.Enabled = False: cmdAllCls.Enabled = False
    zlCommFun.ShowFlash "正在进行门诊费别调整和实收金额计算，请稍后..."
    Screen.MousePointer = 11
    On Error GoTo errHandle
    
    'Zl_门诊划价_Recalcmoney
    strSql = "Zl_门诊划价_Recalcmoney("
    '  病人id_In 门诊费用记录.病人id%Type,
    strSql = strSql & "" & mlng病人ID & ","
    '  费别_In   门诊费用记录.费别%Type,
    strSql = strSql & "'" & Trim(cbo调整费别.Text) & "',"
    '  Nos_In    门诊费用记录.NO%Type := Null
    strSql = strSql & IIf(strNos = "", "NULL,", "'" & strNos & "',")
    '  记录性质_In    门诊费用记录.记录性质 %Type := 1
    strSql = strSql & "1,"
    '   调整费别_In integer:=0
    strSql = strSql & IIf(strNos = "", "0)", "1)")
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    zlCommFun.StopFlash
    MsgBox "调整费别成功!", vbInformation + vbOKOnly, gstrSysName
    
    '移除读入单据时可能加入的无效费别，97338
    For i = cbo调整费别.ListCount - 1 To 0 Step -1
        If Val(cbo调整费别.ItemData(i)) = 0 And cbo调整费别.ListIndex <> i Then
            cbo调整费别.RemoveItem i: Exit For
        End If
    Next
    
    fraOk.Enabled = True: cmdAllSel.Enabled = True: cmdAllCls.Enabled = True
    If strComMand = "所有划价单" Then txt费别.Text = cbo调整费别.Text
    Call cmdFind_Click
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    fraOk.Enabled = True: cmdAllSel.Enabled = True: cmdAllCls.Enabled = True
End Sub

Private Sub chk缺省_Click(Index As Integer)
    Dim i As Long, j As Long
    If mblnNotClick Then Exit Sub
    mblnNotClick = True
    If chk缺省(Index).Value = 1 Then
        For i = 0 To chk缺省.Count - 1
            If i <> Index Then
                chk缺省(i).Value = 0
            End If
        Next
    Else
        j = IIf(Index > 1, Index - 1, chk缺省.Count - 1)
        For i = 0 To chk缺省.Count - 1
            If i <> j Then
                chk缺省(i).Value = 0
            Else
                chk缺省(i).Value = 1
            End If
        Next
    End If
    mblnNotClick = False
    Call ShowBills
End Sub

Private Sub cmdAllCls_Click()
    Call SelBill(True)
End Sub

Private Sub cmdAllSel_Click()
 Call SelBill(False)
End Sub

Private Sub cmdCancel_Click()
    mstrBills = ""
    Unload Me
End Sub

Private Sub cmdFind_Click()
    
    If dtpBegin.Value >= dtpEnd.Value Then
        If Visible Then
            MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
            dtpBegin.SetFocus
        End If
        Exit Sub
    End If
    
    Call ShowBills
    
    If Visible And mshList.Rows > 1 Then
        If mshList.TextMatrix(1, mshList.ColIndex("单据号")) <> "" Then
            mshList.SetFocus
        Else
            txtPatient.SetFocus
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strPati As String, i As Long
    Dim strDept As String
    Dim strNos As String, strNos1 As String
    Dim strNo As String
    Dim cllPro As New Collection
    Dim strSql As String
    
    If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    mstrBills = ""
    If mshList.Rows < 2 Then
        MsgBox "该病人没有任何划价单据可以收费。", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    If mshList.TextMatrix(1, mshList.ColIndex("单据号")) = "" Then
        MsgBox "该病人没有任何划价单据可以收费。", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    
    strNos = "": strNos1 = ""
    For i = 1 To mshList.Rows - 1
        strNo = Trim(mshList.TextMatrix(i, mshList.ColIndex("单据号")))
        If mshList.TextMatrix(i, mshList.ColIndex("选择")) <> "" Then
            If InStr(1, strNos & ",", "," & strNo & ",") = 0 Then
                strNos = strNos & "," & strNo
            End If
            '102748,单据号从小到大
            mstrBills = mshList.TextMatrix(i, mshList.ColIndex("单据号")) & "," & mstrBills
            If InStr(strPati & ",", "," & mshList.TextMatrix(i, mshList.ColIndex("姓名")) & ",") = 0 Then
                strPati = strPati & "," & mshList.TextMatrix(i, mshList.ColIndex("姓名"))
            End If
            If InStr(strDept & ",", "," & mshList.TextMatrix(i, mshList.ColIndex("开单科室")) & ",") = 0 Then
                strDept = strDept & "," & mshList.TextMatrix(i, mshList.ColIndex("开单科室"))
            End If
        Else
            If InStr(1, strNos1 & ",", "," & strNo & ",") = 0 Then
                strNos1 = strNos1 & "," & strNo
            End If
        End If
        If Len(strNos) >= 4000 Then
            strNos = Mid(strNos, 2)
            'Zl_收费划价_暂不执行
            strSql = "Zl_收费划价_暂不执行("
            '  Nos_In      Varchar2,
            strSql = strSql & "'" & strNos & "',"
            '  暂不执行_In Integer:=-1
            strSql = strSql & "0)"
            zlAddArray cllPro, strSql
            strNos = ""
        End If
        If Len(strNos1) >= 4000 Then
            strNos1 = Mid(strNos1, 2)
            'Zl_收费划价_暂不执行
            strSql = "Zl_收费划价_暂不执行("
            '  Nos_In      Varchar2,
            strSql = strSql & "'" & strNos1 & "',"
            '  暂不执行_In Integer:=-1
            strSql = strSql & "-1)"
            zlAddArray cllPro, strSql
            strNos1 = ""
        End If
    Next
    If strNos <> "" Then
         strNos = Mid(strNos, 2)
         'Zl_收费划价_暂不执行
         strSql = "Zl_收费划价_暂不执行("
         '  Nos_In      Varchar2,
         strSql = strSql & "'" & strNos & "',"
         '  暂不执行_In Integer:=-1
         strSql = strSql & "0)"
         zlAddArray cllPro, strSql
         strNos = ""
     End If
     If strNos1 <> "" Then
        strNos1 = Mid(strNos1, 2)
        'Zl_收费划价_暂不执行
        strSql = "Zl_收费划价_暂不执行("
        '  Nos_In      Varchar2,
        strSql = strSql & "'" & strNos1 & "',"
        '  暂不执行_In Integer:=-1
        strSql = strSql & "-1)"
        zlAddArray cllPro, strSql
        strNos1 = ""
    End If
    Err = 0: On Error GoTo ErrHand:
    '先处理划价单:38281
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo ErrHand1:
    
    If mstrBills <> "" Then '102748,单据号从小到大,去掉最后一个分隔符
        mstrBills = Left(mstrBills, Len(mstrBills) - 1)
    End If
    If strPati <> "" Then strPati = Mid(strPati, 2)
    If strDept <> "" Then strDept = Mid(strDept, 2)
    
    If mbln不允许多单据 Then
        If UBound(Split(mstrBills, ",")) > 0 Then
            MsgBox "不允许选择多张划价单收费!", vbInformation, gstrSysName
            mstrBills = ""
            mshList.SetFocus: Exit Sub
        End If
    End If
    
    If mstrBills = "" Then
        MsgBox "请至少选择一张需要收费的划价单据。", vbInformation, gstrSysName
        mstrBills = ""
        mshList.SetFocus: Exit Sub
    '    ElseIf UBound(Split(strDept, ",")) > 0 Then
    '        MsgBox "所选择的多张单据来自不同的开单科室，请分开收费。", vbInformation, gstrSysName
    '        mshList.SetFocus: Exit Sub
    ElseIf UBound(Split(mstrBills, ",")) + 1 >= 200 Then
        MsgBox "单据数量太多，请分成多次收费。", vbInformation, gstrSysName
        mstrBills = ""
        mshList.SetFocus: Exit Sub
    ElseIf UBound(Split(strPati, ",")) > 0 Then
        If MsgBox("选择的单据中包含多个不同的病人姓名，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mstrBills = ""
            mshList.SetFocus: Exit Sub
        End If
    End If
    
    If gblnCheckTest Then
        If Not CheckTest(mstrBills, dtpBegin.Value, dtpEnd.Value) Then
            mstrBills = ""
            mshList.SetFocus: Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
ErrHand:
     gcnOracle.RollbackTrans
ErrHand1:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
 '   If mlng病人ID <> 0 Then mshList.SetFocus
    If cmdOK.Enabled Then cmdOK.SetFocus
End Sub
Private Sub SelBill(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择单据
    '入参:blnCls-是否清除
    '编制:刘兴洪
    '日期:2011-03-16 11:21:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
                 .TextMatrix(i, .ColIndex("选择")) = IIf(blnCls, "", "√")
            End If
        Next
    End With
      stbThis.Panels(2).Text = GetBillNote
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If KeyCode = 13 Then
        If Me.ActiveControl Is mshList Then
            If Me.cmdOK.Enabled And Me.cmdOK.Visible Then
                cmdOK.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
         End If
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call SelBill(False)
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
       Call SelBill(True)
    '54538:刘尔旋,2014-02-24,在选择划价单进行收取时新增快捷键F3的支持
    ElseIf KeyCode = vbKeyF3 Then
        mshList.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Function Init费别() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费别
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-17 21:28:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle

    strSql = _
        "Select a.编码, a.名称, a.简码, Nvl(a.缺省标志, 0) As 缺省, Nvl(a.仅限初诊, 0) As 初诊" & vbNewLine & _
        "From 费别 A" & vbNewLine & _
        "Where Nvl(a.服务对象, 3) In (1, 3) And a.属性 = 1 And Trunc(Sysdate) Between Nvl(a.有效开始, To_Date('1900-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "      And Nvl(a.有效结束, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "Order By a.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    mblnNotClick = True
    With cbo调整费别
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!名称)
            .ItemData(.NewIndex) = 1 '标记有效费别
            If Val(NVL(rsTemp!缺省)) = 1 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
    End With
    mblnNotClick = False
    Init费别 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Form_Load()
    Dim Curdate As Date, blnCancel As Boolean
    Dim i As Integer, j As Long
    
    '选检查主界面中是否发送了回车键的
    mblnCacheKeyReturn = False
    If mblnPreCard Then
        mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    End If
    mlngModule = 1121
    i = Val(zlDatabase.GetPara("缺省选择划价单", glngSys, mlngModule, "1", Array(chk缺省(0), chk缺省(1), chk缺省(2)), InStr(1, mstrPrivs, ";参数设置;") > 0))
    i = IIf(i > 2, 1, i): i = IIf(i < 0, 1, i)
    mblnNotClick = True
    chk缺省(0).Value = 0
    chk缺省(1).Value = 0
    chk缺省(2).Value = 0
    chk缺省(i).Value = 1
    mblnNotClick = False
    RestoreWinState Me, App.ProductName
    If InStr(1, mstrPrivs, ";调整病人费别;") > 0 Then Call Init费别
    cbo调整费别.Visible = InStr(1, mstrPrivs, ";调整病人费别;") > 0
    lbl调整费别.Visible = InStr(1, mstrPrivs, ";调整病人费别;") > 0

    mstrBills = ""
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(Curdate - gintSeekDays, "yyyy-MM-dd HH:mm:ss")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    dtpBegin.MaxDate = Curdate
    dtpEnd.MaxDate = Curdate
    '74296,冉俊明,2014-7-4,按单据的诊断过滤,把单据中的诊断组成一个下拉列表供选择
    cboDiagnose.Clear
    cboDiagnose.AddItem "所有诊断"
    cboDiagnose.ListIndex = cboDiagnose.NewIndex
    
    Call SetHeader
    Call SetDetail
    Call SetActiveList
    If mlng病人ID <> 0 Then
        txtPatient.Locked = True
        txtPatient.TabStop = False
        txtPatient.BackColor = &HE0E0E0
        txt性别.BackColor = &HE0E0E0
        txt年龄.BackColor = &HE0E0E0
        txt费别.BackColor = &HE0E0E0
        txt门诊号.BackColor = &HE0E0E0
        txt付款方式.BackColor = &HE0E0E0
        
        txtPatient.Text = "-" & mlng病人ID
        Call txtPatient_Validate(blnCancel)
        If Not blnCancel Then
            Call cmdFind_Click
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim sngHeight As Single
    
     On Error Resume Next
    fraPati.Left = 20
    fraPati.Top = 50
    fraDays.Top = 50
    sngHeight = stbThis.Height + fraOk.Height
    
    '74296,冉俊明,2014-7-4,按单据的诊断过滤,把单据中的诊断组成一个下拉列表供选择
    lblDiagnose.Top = fraPati.Top + fraPati.Height + 50
    cboDiagnose.Top = lblDiagnose.Top - 40
    
    '59399
    With mshList
         .Left = 0
         .Top = cboDiagnose.Top + cboDiagnose.Height + 20
         .Width = ScaleWidth
         .Height = ScaleHeight - mshList.Top - IIf(mshDetail.Height < 0, 0, mshDetail.Height) - fraHsc.Height - sngHeight
    End With
    
    fraHsc.Left = 0
    fraHsc.Top = mshList.Top + mshList.Height
    fraHsc.Width = ScaleWidth
    
    mshDetail.Left = 0
    mshDetail.Top = fraHsc.Top + fraHsc.Height
    mshDetail.Width = ScaleWidth
    If Me.ScaleHeight - mshDetail.Top - sngHeight <= 1600 Then
        mshDetail.Height = 2000
    Else
        mshDetail.Height = Me.ScaleHeight - mshDetail.Top - sngHeight
    End If
    fraOk.Top = Me.ScaleHeight - sngHeight
    fraOk.Width = Me.ScaleWidth - fraOk.Left
    fraOk.ZOrder
    stbThis.ZOrder
    cmdCancel.Left = fraOk.Width + fraOk.Left - cmdCancel.Width - 50
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng病人ID = 0
    Set mrsList = Nothing
    Set mrsDetail = Nothing
    zl_vsGrid_Para_Save 0, mshList, Me.Caption, "表头列表", False, True
    SaveWinState Me, App.ProductName
    Call zlDatabase.SetPara("缺省选择划价单", IIf(chk缺省(0).Value = 1, 0, IIf(chk缺省(1).Value = 1, 1, 2)), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0)
End Sub

Private Sub fraHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        fraHsc.Top = fraHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub mshDetail_LostFocus()
    Call SetActiveList
End Sub

Private Sub mshList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 0, mshList, Me.Caption, "表头列表", False, True
End Sub

Private Sub mshList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
  zl_vsGrid_Para_Save 0, mshList, Me.Caption, "表头列表", False, True
  With mshList
        If .ColIndex("诊断") >= 0 Then
            .AutoSize .ColIndex("诊断"), .ColIndex("诊断")
        End If
    End With
End Sub
Private Sub mshList_DblClick()
    Call mshList_KeyPress(32)
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, mshList.ColIndex("单据号"))
    If mshList.Row = 0 Or strNo = "" Then Exit Sub
    Call ShowDetail(strNo)
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        If mshList.TextMatrix(mshList.Row, mshList.ColIndex("单据号")) <> "" Then
            If mshList.TextMatrix(mshList.Row, mshList.ColIndex("选择")) = "" Then
                mshList.TextMatrix(mshList.Row, mshList.ColIndex("选择")) = "√"
            Else
                mshList.TextMatrix(mshList.Row, mshList.ColIndex("选择")) = ""
            End If
        End If
        stbThis.Panels(2).Text = GetBillNote
    End If
End Sub

Private Sub mshList_LostFocus()
    Call SetActiveList
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    strHead = "诊断,1,3500|选择,4,500|单据号,4,850|开单科室,1,1200|医生,1,800|姓名,1,800|性别,4,500|年龄,4,500|应收金额,7,850|实收金额,7,850|划价人,1,800|划价时间,4,1850|皮试,4,500"
    With mshList
        .Redraw = flexRDNone
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .RowHeight(0) = 320
        zl_vsGrid_Para_Restore 0, mshList, Me.Caption, "表头列表", False, False
        'If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .Col = 0: .ColSel = .COLS - 1
        If .ColIndex("诊断") >= 0 Then
            .AutoSize .ColIndex("诊断"), .ColIndex("诊断")
        End If
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Integer
    
    strHead = "类别,1,750|项目,1,2000" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,2000", "") & "|规格,1,1000|单位,4,500|数次,7,850|费别,1,750|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,1,850|摘要,1,2000"
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        .RowHeight(0) = 320
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        .Redraw = True
    End With
End Sub

Private Sub SetActiveList(Optional obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &H8000000D
        mshDetail.BackColorSel = &H8000000C
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000D
    Else
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000C
    End If
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.COLS - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub ShowDetail(ByVal strNo As String)
    Dim i As Integer, strSql As String
    
    On Error GoTo errH
    
    strSql = _
    " Select C.名称 as 类别,'['||B.编码||']'||Nvl(E.名称,B.名称) as 项目," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
            IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
    "       Ltrim(To_Char(Avg(Nvl(A.付数,1)*A.数次)" & _
            IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'999990.00000')) as 数次, " & _
    "       A.费别,Ltrim(To_Char(Sum(A.标准单价)" & _
            IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'99999" & gstrFeePrecisionFmt & "')) as 单价, " & _
    "       Ltrim(To_Char(Sum(A.应收金额),'99999" & gstrDec & "')) as 应收金额, " & _
    "       Ltrim(To_Char(Sum(A.实收金额),'99999" & gstrDec & "')) as 实收金额, " & _
    "       D.名称 as 执行科室,A.摘要" & _
    " From 门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
    "       And A.记录性质=1 and A.记录状态 IN(0,1,3) And A.NO=[1]" & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
    " Group by Nvl(A.价格父号,A.序号),C.名称,B.编码,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称,", "") & " B.规格,A.计算单位,A.费别," & _
    "       D.名称,A.摘要,X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
    " Order by Nvl(A.价格父号,A.序号)"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    mshDetail.Clear
    mshDetail.Rows = 2
    If Not mrsDetail.EOF Then
        Set mshDetail.DataSource = mrsDetail
    End If
    Call SetDetail
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ShowBills()
    Dim strSql As String, i As Long
    Dim bytType As Byte
    Dim strTemp As String, strTempCbo As String, sngCboWidth As Long
    
    On Error GoTo errHandle
    bytType = IIf(chk缺省(0).Value = 1, 0, IIf(chk缺省(1).Value = 1, 1, 2))
    strTempCbo = IIf(cboDiagnose.Text = "所有诊断", "", cboDiagnose.Text)
    Screen.MousePointer = 11
    Set mrsList = GetPriceBills(mlng病人ID, mlng挂号科室, dtpBegin.Value, dtpEnd.Value, True, bytType)
    mshList.Clear
    mshList.Rows = 2
    mshDetail.Clear
    mshDetail.Rows = 2
    stbThis.Panels(2).Text = ""
    
    '74296,冉俊明,2014-7-4,按单据的诊断过滤,把单据中的诊断组成一个下拉列表供选择
    If Not mrsList.EOF Then
        mrsList.MoveFirst
        Do While Not mrsList.EOF
            strTemp = NVL(mrsList!诊断)
            mblnNotClick = True
            If zlControl.CboLocate(cboDiagnose, strTemp) = False Then
                cboDiagnose.AddItem strTemp '加入下拉框
            End If
            mblnNotClick = False
            If sngCboWidth < Me.TextWidth(strTemp) Then
                sngCboWidth = Me.TextWidth(strTemp) '下拉文本最大宽度
            End If
            mrsList.MoveNext
        Loop
        '设置下拉框的宽度
        If sngCboWidth + 300 > cboDiagnose.Width Then zlControl.CboSetWidth cboDiagnose.hWnd, sngCboWidth + 300
        
        mrsList.Filter = IIf(strTempCbo = "", "", "诊断='" & strTempCbo & "'")
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = GetBillNote
    End If
    
    Call SetHeader
    Call SetDetail
    Call mshList_EnterCell
    mblnNotClick = True
    If zlControl.CboLocate(cboDiagnose, IIf(strTempCbo = "", "所有诊断", strTempCbo)) = False Then cboDiagnose.ListIndex = 0 '重新定位
    mblnNotClick = False
    
    Me.Refresh
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
End Sub

Private Function GetBillNote() As String
    Dim curTotal As Currency, i As Long, k As Long
    
    k = 0
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, mshList.ColIndex("选择")) <> "" Then
            k = k + 1
            curTotal = curTotal + Val(mshList.TextMatrix(i, mshList.ColIndex("实收金额")))
        End If
    Next
    If k > 0 Then
        GetBillNote = "当前选择了 " & k & " 张单据，合计 " & Format(curTotal, gstrDec) & " 元"
    End If
End Function

Private Sub txtPatient_GotFocus()
    mblnCard = False
    Call zlControl.TxtSelAll(txtPatient)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    mblnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.bln缺省卡号密文)
    '刷卡自动确认
    If mblnCard And Len(txtPatient.Text) = gobjSquare.bln缺省卡号密文 - 1 And KeyAscii <> 8 And KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
        
    If txtPatient.Text <> "" Then
        If txtPatient.Text <> txtPatient.Tag Then
            Set rsTmp = GetPatient(txtPatient.Text, mblnCard)
            If rsTmp Is Nothing Then
                If Visible Then MsgBox "未找到病人的信息。", vbInformation, gstrSysName
                txtPatient.Text = ""
            Else
                
                '就诊卡密码检查
                If Mid(gstrCardPass, 3, 1) = "1" And mblnCard Then
                    If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, rsTmp!姓名, "" & rsTmp!性别, "" & rsTmp!年龄) Then
                        Set rsTmp = Nothing: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                    End If
                End If
            
                txtPatient.PasswordChar = ""
                '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
                txtPatient.IMEMode = 0
                txtPatient.Text = NVL(rsTmp!姓名)
                txtPatient.Tag = txtPatient.Text
                mlng病人ID = NVL(rsTmp!病人ID, 0)
                txt性别.Text = NVL(rsTmp!性别)
                txt年龄.Text = NVL(rsTmp!年龄)
                txt费别.Text = NVL(rsTmp!费别)
                txt门诊号.Text = NVL(rsTmp!门诊号)
                txt付款方式.Text = NVL(rsTmp!医疗付款方式)
                If InStr(1, mstrPrivs, ";调整病人费别;") > 0 Then
                    mblnNotClick = True
                    Local费别 Trim(txt费别.Text), True
                    mblnNotClick = False
                End If
            End If
        End If
    End If
    If txtPatient.Text = "" Then
        mlng病人ID = 0
        txtPatient.Tag = ""
        txt性别.Text = ""
        txt年龄.Text = ""
        txt费别.Text = ""
        txt门诊号.Text = ""
        txt付款方式.Text = ""
        Cancel = True: Exit Sub
    End If
End Sub

Private Function GetPatient(ByVal strInput As String, ByVal blnCard As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 16:47:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strWhere As String
    Dim strPati As String, vRect As RECT
    
    strInput = UCase(strInput)
    
    '病人输入的权限
    If gint病人来源 = 1 Then
        'strWhere = " And Nvl(A.当前科室ID,0)=0"
        If Not mbln住院病人门诊收费 Then    '34182
            strWhere = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
        End If
    ElseIf gint病人来源 = 2 Then
        If Not mbln住院病人门诊收费 Then    '34182
            strWhere = " And Nvl(A.当前科室ID,0)<>0"
        End If
    End If
    
    '读取病人信息
    strSql = "Select A.病人ID,A.门诊号,A.姓名,A.性别,A.年龄,A.就诊卡号,A.卡验证码,A.费别,A.医疗付款方式,A.病人类型,A.险类 From 病人信息 A Where 1=1"
    If blnCard Then '就诊卡号
        If gint病人来源 = 1 And Not gblnInputCard Then Exit Function
        '见问题:27364
        If Not gobjSquare.objDefaultCard Is Nothing And gobjSquare.bln按缺省卡查找 Then
            lng卡类别ID = gobjSquare.objDefaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSql = strSql & strWhere & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSql = strSql & strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSql = strSql & strWhere & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strSql = strSql & strWhere & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "." Then '挂号单号
        If gint病人来源 = 1 And Not gblnInputNO Then Exit Function
        '按日或年顺序编号规则
        strInput = GetFullNO(Mid(strInput, 2), 12)
        strSql = "" & _
        " Select A.病人ID,Nvl(A.门诊号,B.标识号) as 门诊号,A.病人类型,A.险类," & _
        "       Nvl(A.姓名,B.姓名) as 姓名,Nvl(A.性别,B.性别) as 性别," & _
        "       Nvl(A.年龄,B.年龄) as 年龄,A.就诊卡号,A.卡验证码,Nvl(A.费别,B.费别) as 费别," & _
        "       Nvl(A.医疗付款方式,C.名称) as 医疗付款方式" & _
        " From 病人信息 A,门诊费用记录 B,医疗付款方式 C" & _
        " Where B.记录性质=4 And B.记录状态=1 And B.NO=[2]" & _
        "       And B.病人ID=A.病人ID(+) And B.付款方式=C.编码(+)" & strWhere & _
             zlGetRegEventsCons("加班标志", "B")
    Else
        '通过姓名模糊查找病人
        If gblnSeekName Then
            strWhere = " A.姓名 Like '" & strInput & "%' " & strWhere
            strPati = _
                " Select A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.就诊卡号,A.卡验证码,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                " From 病人信息 A Where " & strWhere & _
                IIf(gintNameDays = 0, "", " And (A.就诊时间>Trunc(Sysdate-" & gintNameDays & ") Or A.登记时间>Trunc(Sysdate-" & gintNameDays & "))") & _
                " And Rownum<101" & _
                " Order by A.姓名"
            vRect = zlControl.GetControlRect(txtPatient.hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strPati, 0, "病人Find", , , , , , True, vRect.Left, vRect.Top, txtPatient.Height, , , True)
            If rsTmp Is Nothing Then Exit Function
            If rsTmp.EOF Then Exit Function
            strInput = rsTmp!病人ID
            strSql = strSql & strWhere & " And A.病人ID=[2]"
        Else
            Exit Function
        End If
    End If
        
    On Error GoTo errH
    '75259:李南春,2014-7-10，病人姓名显示颜色处理
    txtPatient.ForeColor = Me.ForeColor
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If Not rsTmp.EOF Then
        Call SetPatiColor(txtPatient, NVL(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), txtPatient.ForeColor, vbRed))
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = NVL(rsTmp!卡验证码)
        Set GetPatient = rsTmp
    End If
    Exit Function
NotFoundPati:
    Set GetPatient = Nothing
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
