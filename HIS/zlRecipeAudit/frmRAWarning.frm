VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRAWarning 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动审查提醒"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frmRAWarning.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "忽略提醒(&I)"
      Height          =   360
      Left            =   5880
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   360
      Left            =   7320
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame fraUnqualified 
      Caption         =   "不合格明细"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VSFlex8Ctl.VSFlexGrid vsfUnqualified 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8415
         _cx             =   14843
         _cy             =   5106
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
         FixedCols       =   0
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
End
Attribute VB_Name = "frmRAWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnAdjust As Boolean                   'True调整药品；False忽略提醒

Private Const MSTR_VSF As String = "编码,,3,1000|审查项目,,3,2500|药品名称,,3,3000|规格,,3,1500"

Public Function ShowMe(ByVal strNG As String, ByVal frmOwner As Form) As Boolean
'功能：显示窗体接口方法
'参数：
'  strInfo：不合格的项目和医嘱；  格式：项目ID,医嘱ID[|项目ID,医嘱ID]
'  frmOwner：宿主窗体对象
'返回：True操作调整药品；False操作忽略提醒

    Dim rsUnqualified As ADODB.Recordset
    
    mblnAdjust = True   '默认调整药品
    
    '加载数据
    InitVSF vsfUnqualified
    mdlDefine.SetVSFHead vsfUnqualified, MSTR_VSF
    
    If GetData(strNG, rsUnqualified) = True Then
        mdlDefine.FillVSFData vsfUnqualified, rsUnqualified
    End If
    
    '显示窗体
    Show vbModal, frmOwner
    
    '返回操作值
    ShowMe = mblnAdjust
    
End Function

Private Function GetData(ByVal strNG As String, ByRef rsVar As ADODB.Recordset) As Boolean
    Dim strSQL As String
    
    strSQL = "Select b.名称 药品名称, b.规格, d.编码, d.简称 审查项目 " & vbCr & _
             "From 病人医嘱记录 A, 收费项目目录 B, 处方审查项目 D, Table(f_Num2list2([1], '|', ',')) E " & vbCr & _
             "Where e.C1 = d.Id And e.C2 = a.Id(+) And b.Id(+) = a.收费细目id " & vbCr & _
             "Order By d.编码, b.名称"
             
    On Error GoTo errHandle
    Set rsVar = zlDatabase.OpenSQLRecord(strSQL, "获取不合格审查项目", strNG)
    
    GetData = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdIgnore_Click()
    mblnAdjust = False
    Unload Me
End Sub

Private Sub cmdClose_Click()
    mblnAdjust = True
    Unload Me
End Sub

Private Sub InitVSF(ByVal vsfVar As VSFlexGrid)
'功能：初始化窗体的VSFlexGrid控件的风格
'参数：
'  vsfVar：要初始化的VSFlexGrid控件

    With vsfVar
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
    End With
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
End Sub
