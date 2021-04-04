VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditAssistant 
   AutoRedraw      =   -1  'True
   Caption         =   "词句选择"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10950
   Icon            =   "frmEditAssistant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDef 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   2955
      Picture         =   "frmEditAssistant.frx":058A
      ScaleHeight     =   3780
      ScaleWidth      =   5790
      TabIndex        =   17
      Top             =   1455
      Visible         =   0   'False
      Width           =   5790
      Begin VB.CommandButton cmdAdd 
         Height          =   270
         Left            =   5280
         Picture         =   "frmEditAssistant.frx":47BAC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "加入字段(ALT+A)"
         Top             =   2745
         Width           =   270
      End
      Begin VB.ComboBox cbo字段 
         Height          =   300
         Left            =   1125
         TabIndex        =   24
         Top             =   2730
         Width           =   4125
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "检查(&K)"
         Height          =   350
         Left            =   2175
         TabIndex        =   23
         Top             =   3285
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4470
         TabIndex        =   22
         Top             =   3285
         Width           =   1100
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3375
         TabIndex        =   21
         Top             =   3285
         Width           =   1100
      End
      Begin VB.TextBox txtAdvice 
         Height          =   1125
         Left            =   1125
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1575
         Width           =   4440
      End
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1245
         Width           =   4440
      End
      Begin VB.PictureBox picTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1545
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   90
         Width           =   240
         Begin VB.Image imgTip 
            Height          =   240
            Left            =   0
            Picture         =   "frmEditAssistant.frx":47C76
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Label lblDefTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "导入内容自定义"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   105
         Width           =   1365
      End
      Begin VB.Image imgClose 
         Height          =   285
         Left            =   5445
         Picture         =   "frmEditAssistant.frx":4E4C8
         Stretch         =   -1  'True
         Top             =   45
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   30
         X2              =   6090
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   -60
         X2              =   6000
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   105
         X2              =   6165
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   15
         X2              =   6075
         Y1              =   3165
         Y2              =   3165
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditAssistant.frx":4E932
         Height          =   645
         Left            =   450
         TabIndex        =   29
         Top             =   495
         Width           =   5040
      End
      Begin VB.Label lbl类别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类    别"
         Height          =   180
         Left            =   330
         TabIndex        =   28
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lbl内容格式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内容格式"
         Height          =   180
         Left            =   330
         TabIndex        =   27
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl字段项目 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "字段项目"
         Height          =   180
         Left            =   330
         TabIndex        =   26
         Top             =   2790
         Width           =   1455
      End
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   3465
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   3465
      MousePointer    =   7  'Size N S
      TabIndex        =   12
      Top             =   3150
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   3375
      TabIndex        =   11
      Top             =   2865
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   4125
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   10950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6525
      Width           =   10950
      Begin VB.CommandButton cmdDef 
         Caption         =   "导入内容自定义(&F)"
         Height          =   350
         Left            =   6030
         TabIndex        =   31
         Top             =   150
         Width           =   1740
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "定位(&L)"
         Height          =   350
         Left            =   2715
         TabIndex        =   7
         Top             =   135
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   870
         TabIndex        =   6
         Top             =   135
         Width           =   1845
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9375
         TabIndex        =   9
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8160
         TabIndex        =   8
         Top             =   135
         Width           =   1100
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "请输入查找条件"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   3975
         TabIndex        =   16
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "词句查找"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Frame fraUD 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3465
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   3765
      Width           =   5475
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "详细内容"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   30
         Width           =   720
      End
   End
   Begin RichTextLib.RichTextBox rtfSentence 
      Height          =   1245
      Left            =   3540
      TabIndex        =   2
      Top             =   4680
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   2196
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmEditAssistant.frx":4E9DD
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   3285
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2400
      Left            =   3390
      TabIndex        =   1
      Top             =   225
      Width           =   6315
      _cx             =   11139
      _cy             =   4233
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmEditAssistant.frx":4EA7A
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
      ExplorerBar     =   5
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
      Begin MSComctlLib.ImageList imgList 
         Left            =   420
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":4EAEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":4F089
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":4F623
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditAssistant.frx":4FBBD
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditAssistant.frx":50157
            Key             =   "Expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5865
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   10345
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line lin 
      Index           =   1
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line lin 
      Index           =   2
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line lin 
      Index           =   3
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Line lin 
      Index           =   4
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Line lin 
      Index           =   5
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line lin 
      Index           =   6
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line lin 
      Index           =   7
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3165
      Y2              =   3165
   End
End
Attribute VB_Name = "frmEditAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'===============================================================================================
Public mblnShow As Boolean '该窗体是否正在显示
Private mstrInput As String
Private mstrSentence As String
Private mstrLike As String
Private mintType As Integer
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mint婴儿 As Integer
Private mblnOK As Boolean

Private mlngPreY As Long

Private mrsPati As New ADODB.Recordset
Private mrsFind As New ADODB.Recordset
Private mrsField As ADODB.Recordset
Private mrsFormat As ADODB.Recordset
Private mobjPublicLis As Object
Private mobjXML As Object
Private mstrXmlVersion As String
Private mintIndex As Integer
Private mobjVBA As Object
Private mobjScript As clsScript

Private Type LisItem
    检验报告id As String
    申请id As String
    紧急标志 As Integer
    检验项目 As String
    标本序号 As String
    是否微生物 As Integer
    检验次数 As Integer
    检验人 As String
    审核人 As String
    审核时间 As String
    申请时间 As String
End Type

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, Optional ByVal strInput As String, Optional ByVal intType As Integer = 3) As String
    mstrSentence = ""
    mstrInput = strInput
    mintType = intType
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mint婴儿 = int婴儿
    
    On Error Resume Next
    Me.Show 1, frmParent
    Err.Clear: On Error GoTo 0
    
    If mblnOK Then
        ShowMe = mstrSentence
    Else
        ShowMe = mstrInput
    End If
End Function

Private Function ShowTree() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strMatch As String
    Dim strXMLLIS As String
    Dim objXMLNodeList As Object, objXMLNode As Object
    Dim lngParentID As Long, lngID As Long
    Dim strFirstName As String
    Dim rsItem As New ADODB.Recordset
    Dim rsLisItem As ADODB.Recordset
    Dim L_Item As LisItem
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1"
        
    '98483:刘鹏飞,2016-11-30,性能优化
    strSQL = _
        " Select Max(Level) As 级数, a.Id, a.上级id, a.编码, a.名称, a.说明, Max(b.分类id) 分类id" & vbNewLine & _
        " From 病历词句分类 a," & vbNewLine & _
        "     (Select 分类id" & vbNewLine & _
        "       From (Select a.Id, a.分类id" & vbNewLine & _
        "              From 病历词句分类 b, 病历词句示范 a" & vbNewLine & _
        "              Where a.分类id = b.Id And Nvl(Substr(b.范围, [1], 1), '0') = '1' And" & vbNewLine & _
        "                    ((Nvl(a.通用级, 0) = 0 Or a.通用级 = 1 And a.科室id In (Select a.部门id From 部门人员 a Where a.人员id = [11]) Or" & vbNewLine & _
        "                    a.通用级 = 2 And a.人员id = [11])))" & vbNewLine & _
        "       Where " & strMatch & vbNewLine & _
        "       Group By 分类id) b" & vbNewLine & _
        " Where a.Id = b.分类id(+)" & vbNewLine & _
        " Start With a.Id In (b.分类id)" & vbNewLine & _
        " Connect By Prior a.上级id = a.Id" & vbNewLine & _
        " Group By a.Id, a.上级id, a.编码, a.名称, a.说明" & vbNewLine & _
        " Order By 级数 Desc, 编码"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
        CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "", glngUserId)
    
    '添加词句分类
    tvw_s.Nodes.Clear
    Set objNode = tvw_s.Nodes.Add(, , "_", "所有词句", "Close")
    objNode.ExpandedImage = "Expend"
    objNode.Expanded = True
    Do While Not rsTmp.EOF
        Set objNode = tvw_s.Nodes.Add("_" & NVL(rsTmp!上级ID), tvwChild, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Close")
        objNode.Tag = NVL(rsTmp!分类id, 0)
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Loop

    '强制添加医嘱相关结点
    Set objNode = tvw_s.Nodes.Add(, , "=", "所有医嘱", "Close")
    objNode.ExpandedImage = "Expend"
    objNode.Expanded = True
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=1", "输液类", "Close")
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=2", "注射类", "Close")
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=4", "口服类", "Close")
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=0", "其他类", "Close")
    objNode.ExpandedImage = "Expend"
    '120692:添加检验项目
    If Not mobjPublicLis Is Nothing Then
        Call Record_Init(rsLisItem, "id," & adBigInt & ",18|parent_id," & adBigInt & ",18|node_name, " & adVarChar & ",50|node_value," & adVarChar & ",4000")
        Set objNode = tvw_s.Nodes.Add(, , "％", "检验项目", "Close")
        objNode.ExpandedImage = "Expend"
        objNode.Expanded = True
        strXMLLIS = mobjPublicLis.GetLaboratoryReportList(mlng病人ID, mlng主页ID)
        If strXMLLIS <> "" Then
            If OpenXMLDocument(strXMLLIS) = True Then
                'LIS返回的XML信息
'                <检验报告列表>
'                    <检验报告id>54603972</检验报告id>
'                    <申请id>7199230</申请id>
'                    <紧急标志>0</紧急标志>
'                    <检验项目>血型 血常规23项(抗凝血)</检验项目>
'                    <标本序号>7199229</标本序号>
'                    <是否微生物>0</是否微生物>
'                    <检验次数>0</检验次数>
'                    <检验人>杨笑琼</检验人>
'                    <审核人>贾建</审核人>
'                    <审核时间>2008/3/15 17:13:15</审核时间>
'                    <申请时间>2008/3/15 15:29:00</申请时间>
'                    <检验报告id>56511459</检验报告id>
'                    ……
'                <检验报告列表>
                Set objXMLNodeList = mobjXML.selectNodes(".//检验报告列表").Item(0).childNodes
                strFirstName = objXMLNodeList.Item(0).nodename
                lngID = 0
                For Each objXMLNode In objXMLNodeList
                    lngID = lngID + 1
                    If objXMLNode.nodename = strFirstName Then '每次以检验报告ID来区分一个项目
                        lngParentID = lngID
                        rsLisItem.AddNew
                        rsLisItem!ID = lngID
                        rsLisItem!parent_id = 0
                        rsLisItem!node_name = objXMLNode.nodename
                        rsLisItem!node_value = objXMLNode.Text
                        rsLisItem.Update
                        lngID = lngID + 1
                    End If
                    rsLisItem.AddNew
                    rsLisItem!ID = lngID
                    rsLisItem!parent_id = lngParentID
                    rsLisItem!node_name = objXMLNode.nodename
                    rsLisItem!node_value = objXMLNode.Text
                    rsLisItem.Update
                Next
            Else
                MsgBox "LIS接口返回的检验结果XML格式不正确，不能加载检验结果信息！", vbInformation, gstrSysName
            End If
            Set rsItem = zlDatabase.CopyNewRec(rsLisItem)
            rsLisItem.Filter = "parent_id=0"
            
            lngID = 1
            Do While Not rsLisItem.EOF
                rsItem.Filter = "parent_id=" & rsLisItem!ID
                Do While Not rsItem.EOF
                    Select Case rsItem!node_name & ""
                        Case "检验报告id"
                            L_Item.检验报告id = rsItem!node_value & ""
                        Case "申请id"
                            L_Item.申请id = rsItem!node_value & ""
                        Case "紧急标志"
                            L_Item.紧急标志 = Val(rsItem!node_value & "")
                        Case "检验项目"
                            L_Item.检验项目 = rsItem!node_value & ""
                        Case "标本序号"
                            L_Item.标本序号 = rsItem!node_value & ""
                        Case "是否微生物"
                            L_Item.是否微生物 = Val(rsItem!node_value & "")
                        Case "检验次数"
                            L_Item.检验次数 = Val(rsItem!node_value & "")
                        Case "检验人"
                            L_Item.检验人 = rsItem!node_value & ""
                        Case "审核人"
                            L_Item.审核人 = rsItem!node_value & ""
                        Case "审核时间"
                            L_Item.审核时间 = Format(rsItem!node_value & "", "YYYY-MM-DD HH:mm:SS")
                        Case "申请时间"
                            L_Item.申请时间 = Format(rsItem!node_value & "", "YYYY-MM-DD HH:mm:SS")
                    End Select
                    rsItem.MoveNext
                Loop
                lngID = lngID + 1
                Set objNode = tvw_s.Nodes.Add("％", tvwChild, "％" & L_Item.检验报告id & "_" & lngID, L_Item.检验项目 & "[" & L_Item.审核时间 & "]", "Close")
                objNode.Tag = L_Item.检验报告id & "'" & L_Item.申请id & "'" & L_Item.紧急标志 & "'" & L_Item.标本序号 & "'" & L_Item.是否微生物 & "'" & _
                        L_Item.检验次数 & "'" & L_Item.检验人 & "'" & L_Item.审核人 & "'" & L_Item.审核时间 & "'" & L_Item.申请时间
                objNode.ExpandedImage = "Expend"
                rsLisItem.MoveNext
            Loop
        End If
    End If
    
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Selected = True
    End If
    If Not tvw_s.SelectedItem Is Nothing Then
        tvw_s.SelectedItem.Expanded = True
        tvw_s.SelectedItem.EnsureVisible
    End If
    
    Screen.MousePointer = 0
    ShowTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowList(Optional ByVal lng分类id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim int执行分类 As Integer
    Dim strSQL As String, i As Long
    Dim strMatch As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If Mid(tvw_s.SelectedItem.Key, 1, 1) = "_" Then
        Call InitVsf(0)
        strMatch = "f_Sentence_Matched(A.ID,[2],[3],[4],[5],[6],[7],[8],[9],[10],[11])=1"
        If lng分类id <> 0 Then
            '按树形读取数据
            strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Trim(B.内容文本) as 内容文本" & _
                " From 病历词句组成 B,病历词句示范 A" & _
                " Where A.ID=B.词句ID(+) And B.排列次序(+)=1 And A.分类ID=[1] And " & strMatch & _
                "   And ((Nvl(A.通用级, 0) = 0" & _
                "       Or A.通用级 = 1 And A.科室id In (Select A.部门id From 部门人员 A Where A.人员id =[12])" & _
                "       Or A.通用级 = 2 And A.人员id =[12])) Order by A.编号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng分类id, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
                CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "", glngUserId)
        Else
            '按输入读取数据
            strSQL = "Select A.ID,A.编号,A.名称,A.通用级,LPad(B.排列次序,3,'0')||Trim(B.内容文本) as 内容文本" & _
                " From 病历词句分类 C,病历词句组成 B,病历词句示范 A" & _
                " Where A.ID=B.词句ID And Nvl(B.内容性质,0)=0 And A.分类ID=C.ID And Nvl(Substr(C.范围, [1], 1), '0') = '1'" & _
                "   And (A.编号 Like [1]||'%'" & _
                "       Or A.名称 Like " & IIF(mstrLike <> "", "'%'||", "") & "[1]||'%'" & _
                "       Or B.内容文本 Like " & IIF(mstrLike <> "", "'%'||", "") & "[1]||'%')" & _
                "   And ((Nvl(A.通用级, 0) = 0" & _
                "       Or A.通用级 = 1 And A.科室id In(Select A.部门id From 部门人员 A Where A.人员id = [12])" & _
                "       Or A.通用级 = 2 And A.人员id =[12]))"
            
            strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Substr(Min(A.内容文本),4) as 内容文本" & _
                " From (" & strSQL & ") A Where " & strMatch & " Group by A.ID,A.编号,A.名称,A.通用级 Order by A.编号"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstrInput, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
                CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "", glngUserId)
        End If
        vsList.Redraw = flexRDNone
        vsList.Rows = vsList.FixedRows
        If rsTmp Is Nothing Then Screen.MousePointer = 0: Exit Function
        If Not rsTmp.EOF Then
            vsList.Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                vsList.RowData(i) = Val(rsTmp!ID)
                vsList.TextMatrix(i, 1) = NVL(rsTmp!编号)
                vsList.TextMatrix(i, 2) = NVL(rsTmp!名称)
                vsList.TextMatrix(i, 3) = NVL(rsTmp!内容文本)
                vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(NVL(rsTmp!通用级, 0) + 1).Picture
                rsTmp.MoveNext
            Next
            vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
            vsList.ROW = 1: vsList.COL = 2
        End If
        vsList.Redraw = flexRDDirect
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "=" Then
        Call InitVsf(1)
        '91329:提取医嘱包含:给药方式由本科室执行和病区执行"C.执行性质 IN (1,2)"
        '125170,18-07-24,CL,诊疗项目目录的执行科室比病人医嘱记录的执行性质准确
        int执行分类 = lng分类id
        If tvw_s.SelectedItem.Key = "=" Then int执行分类 = 99
        strSQL = "" & _
            " Select a.Id, a.相关id, b.名称 诊疗项目, Decode(Substr('' || Nvl(a.总给予量, 0), 1, 1), '.', 0, '') || a.总给予量 as 总给予量, Decode(Substr('' || Nvl(a.单次用量, 0), 1, 1), '.', 0, '') || a.单次用量 as 单次用量, b.计算单位, a.医嘱内容, a.医生嘱托, a.开嘱医生, a.开始执行时间, d.名称 给药途径, d.执行分类, 1 As 通用级," & vbNewLine & _
            "       a.医嘱内容 || Decode(Substr('' || Nvl(a.单次用量, 0), 1, 1), '.', 0, '') || a.单次用量 || b.计算单位 || c.医嘱内容 As 内容文本" & vbNewLine & _
            " From 病人医嘱记录 a, 诊疗项目目录 b, 病人医嘱记录 c, 诊疗项目目录 d" & vbNewLine & _
            " Where a.诊疗类别 In ('5', '6', '7') And a.诊疗项目id = b.Id And a.病人id = [1] And a.主页id = [2] And a.婴儿 = [3] And c.诊疗类别 = 'E' And" & vbNewLine & _
            "      d.执行科室 In (1, 2, 3, 4, 6) And Nvl(d.执行分类, 0) = [4] And d.Id = c.诊疗项目id And a.相关id = c.Id And c.上次执行时间 Is Not Null" & vbNewLine & _
            " Order By a.开始执行时间 Desc"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID, mint婴儿, int执行分类)
        vsList.Redraw = flexRDNone
        vsList.Rows = vsList.FixedRows
        If rsTmp Is Nothing Then Screen.MousePointer = 0: Exit Function
        If Not rsTmp.EOF Then
            vsList.Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                vsList.RowData(i) = Val(rsTmp!ID)
                vsList.TextMatrix(i, vsList.ColIndex("相关ID")) = NVL(rsTmp!相关ID)
                vsList.TextMatrix(i, vsList.ColIndex("一组")) = ""
                vsList.TextMatrix(i, vsList.ColIndex("医嘱内容")) = NVL(rsTmp!医嘱内容)
                vsList.TextMatrix(i, vsList.ColIndex("单次用量")) = NVL(rsTmp!单次用量) & NVL(rsTmp!计算单位)
                vsList.TextMatrix(i, vsList.ColIndex("给药途径")) = NVL(rsTmp!给药途径)
                vsList.TextMatrix(i, vsList.ColIndex("总给予量")) = NVL(rsTmp!总给予量)
                vsList.TextMatrix(i, vsList.ColIndex("诊疗项目")) = NVL(rsTmp!诊疗项目)
                vsList.TextMatrix(i, vsList.ColIndex("开嘱医生")) = NVL(rsTmp!开嘱医生)
                vsList.TextMatrix(i, vsList.ColIndex("开始时间")) = Format(NVL(rsTmp!开始执行时间), "YYYY-MM-DD HH:mm")
                vsList.TextMatrix(i, vsList.ColIndex("医生嘱托")) = NVL(rsTmp!医生嘱托)
                vsList.TextMatrix(i, vsList.ColIndex("内容文本")) = NVL(rsTmp!内容文本)
                vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(NVL(rsTmp!通用级, 0) + 1).Picture
                
                rsTmp.MoveNext
            Next
            vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
            vsList.ROW = 1: vsList.COL = 2
        End If
        vsList.Redraw = flexRDDirect
        For i = vsList.FixedRows To vsList.Rows - 1
            If vsList.TextMatrix(i, vsList.ColIndex("一组")) = "" Then
                Call SetTag一并给药(i)
            End If
        Next
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "％" Then
        '提取检验结果
        Call ShowLisList
    End If
    Screen.MousePointer = 0
    ShowList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cbo类别_Click()
    Dim arrField As Variant, i As Long
    
    '1.检查并更新当前类别的内容
    '------------------------------
    If Visible Then
        If Not UpdateFormat Then Exit Sub
    End If
    '2.显示新切换到的类别的内容
    '------------------------------
    mintIndex = cbo类别.ListIndex
    
    '显示可用字段列表
    cbo字段.Clear
    mrsField.Filter = "类别=" & cbo类别.ItemData(cbo类别.ListIndex)
    arrField = Split(mrsField!字段, ",")
    For i = 0 To UBound(arrField)
        cbo字段.AddItem arrField(i)
    Next
    
    '显示当前设置的医嘱内容
    mrsFormat.Filter = "类别=" & cbo类别.ItemData(cbo类别.ListIndex)
    If Not mrsFormat.EOF Then
        If Val("" & mrsFormat!是否修改) = 1 Then
            txtAdvice.Text = mrsFormat!新格式 & ""
        Else
            txtAdvice.Text = mrsFormat!格式 & ""
        End If
    Else
        txtAdvice.Text = ""
    End If
    txtAdvice.Tag = ""
End Sub

Private Function UpdateFormat() As Boolean
    Dim strMsg As String
    
    strMsg = CheckFormat(txtAdvice.Text)
    If strMsg <> "" Then
        Call zlControl.CboSetIndex(cbo类别.hWnd, mintIndex)
        MsgBox strMsg, vbInformation, gstrSysName
        txtAdvice.SetFocus: Exit Function
    End If
    mrsFormat.Filter = "类别=" & cbo类别.ItemData(mintIndex)
    If mrsFormat.EOF Then
        If Trim(txtAdvice.Text) <> "" Then '原本没内容的情况下
            mrsFormat.AddNew
            mrsFormat!类别 = cbo类别.ItemData(mintIndex)
            mrsFormat!名称 = cbo类别.List(mintIndex)
            mrsFormat!新格式 = txtAdvice.Text
            mrsFormat!是否修改 = 1
            mrsFormat.Update
        End If
    Else
        If mrsFormat!格式 & "" <> txtAdvice.Text Then
            mrsFormat!新格式 = txtAdvice.Text
            mrsFormat!名称 = cbo类别.List(mintIndex)
            mrsFormat!是否修改 = 1
            mrsFormat.Update
        Else
            If Val(mrsFormat!是否修改 & "") = 1 Then
                mrsFormat!是否修改 = 0
                mrsFormat.Update
            End If
        End If
    End If
    txtAdvice.Tag = ""
    UpdateFormat = True
End Function

Private Function CheckFormat(ByVal strText As String) As String
'功能：检查医嘱内容是否正确
'返回：错误信息
'      strPreview=预览医嘱内容效果
    Dim intLeft As Integer, intRight As Integer
    Dim strTmp As String, strPar As String
    Dim strMsg As String, i As Long
    Dim objVBA As Object, strEval As String
    Dim objScript As New clsScript
    
    If Trim(strText) = "" And strText = Trim(strText) Then Exit Function
    If zlCommFun.ActualLen(strText) > txtAdvice.MaxLength Then
        strMsg = "定义内容太长，只允许 " & txtAdvice.MaxLength & " 个字符或 " & txtAdvice.MaxLength \ 2 & " 个汉字。"
        GoTo EndLine
    End If
    
    If Not InStr(strText, "[") > 0 Then
        strMsg = "格式不正确,须绑定字段项目。"
        GoTo EndLine
    End If
        
    '检查配对情况
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = "[" Then
            intLeft = intLeft + 1
        ElseIf Mid(strText, i, 1) = "]" Then
            intRight = intRight + 1
            If intLeft <> intRight Then
                strMsg = """[""与""]""括号不配对。"
                GoTo EndLine
            End If
        End If
    Next
    If intLeft = 0 And intRight = 0 Then Exit Function
    If intLeft <> intRight Then
        strMsg = """[""与""]""括号不配对。"
        GoTo EndLine
    End If
    
    '检查字段名称
    strTmp = strText
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Trim(Left(strTmp, InStr(strTmp, "]") - 1))
                        
        If strPar = "" Then
            strMsg = """[]""括号之中没有书写字段名。"
            GoTo EndLine
        End If
        
        For i = 0 To cbo字段.ListCount - 1
            If cbo字段.List(i) = "[" & strPar & "]" Then Exit For
        Next
        If i > cbo字段.ListCount - 1 Then
            strMsg = "使用了不存在的""[" & strPar & "]""字段。"
            GoTo EndLine
        End If
    Loop
    
    '执行测试
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    If objVBA Is Nothing Then
        strMsg = "Microsoft Script Control 未正确安装(msscript.ocx)，不能执行检查。请重新安装客户端程序。"
        GoTo EndLine
    End If
    Err.Clear: On Error GoTo 0
    objVBA.Language = "VBScript"
    objVBA.addObject "clsScript", objScript, True
    strEval = Replace(strText, "[", """")
    strEval = Replace(strEval, "]", """")
    On Error Resume Next
    Call objVBA.Eval(strEval)
    If objVBA.Error.Number <> 0 Then
        strMsg = objVBA.Error.Description
        objVBA.Error.Clear
    End If
EndLine:
    CheckFormat = strMsg
End Function

Private Sub cbo字段_GotFocus()
    Call zlControl.TxtSelAll(cbo字段)
End Sub

Private Sub cbo字段_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    If cbo字段.Text = "" Then Exit Sub
    txtAdvice.SelText = cbo字段.Text
    cbo字段.SetFocus
End Sub

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim strMsg As String
    
    If Trim(txtAdvice.Text) <> "" Then
        strMsg = CheckFormat(txtAdvice.Text)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            txtAdvice.SetFocus
        Else
            MsgBox "内容格式书写正确。", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Dim blnCancel As Boolean
    If Not mrsFormat Is Nothing Then
        mrsFormat.Filter = "是否修改=1"
        blnCancel = mrsFormat.RecordCount > 0
        If blnCancel = False Then
            mrsFormat.Filter = "类别=" & cbo类别.ItemData(cbo类别.ListIndex)
            If Not mrsFormat.EOF Then
                blnCancel = (mrsFormat!格式 & "" <> txtAdvice.Text)
            Else
                blnCancel = txtAdvice.Text <> ""
            End If
        End If
    End If
    If blnCancel = True Then
        If MsgBox("如果退出将会丢失你所改变的内容，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            '恢复之前的修改
            If Not mrsFormat Is Nothing Then
                mrsFormat.Filter = "是否修改=1"
                Do While Not mrsFormat.EOF
                    mrsFormat!是否修改 = 0
                    mrsFormat!新格式 = ""
                    mrsFormat.Update
                mrsFormat.MoveNext
                Loop
            End If
            txtAdvice.Text = ""
        End If
    End If
     picDef.Visible = False
     SetControlEnable True
End Sub

Private Sub cmdDef_Click()
    With picDef
        .Left = (Me.ScaleWidth - .Width) \ 2
        .Top = (Me.ScaleHeight - .Height) \ 2
        .Visible = True
        .ZOrder 0
    End With
    SetControlEnable False
End Sub

Private Sub cmdFind_Click()
'功能:词句查找
    Dim strText As String, strMatch As String
    Dim strFind As String, strSQL As String
    Dim lngRow As Long, lngTypeID As Long
    
    On Error GoTo ErrHand
    
    If mrsFind.State = adStateOpen Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocaItem
        Exit Sub
    End If
    
    If Trim(txtFind.Text) = "" Then
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txtFind.Text, "'") > 0 Then
        MsgBox "输入的内容包含非法字符 ' ,请检查!", vbInformation, gstrSysName
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If Not tvw_s.SelectedItem Is Nothing Then
        lngTypeID = Val(tvw_s.SelectedItem.Tag)
    Else
        lngTypeID = 0
    End If
    
    strText = mstrLike & txtFind.Text & "%"
    If zlCommFun.IsCharChinese(txtFind.Text) Then
        strFind = " And A.名称 Like '" & strText & "'"
    ElseIf IsNumeric(txtFind.Text) Then
        strFind = " And A.编号 Like '" & strText & "'"
    Else
        strFind = " And zlspellcode(A.名称) Like '" & UCase(strText) & "'"
    End If
    
    '根据输入的内容提取匹配的词句
    strMatch = " f_Sentence_Matched(A.ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1 "
    strSQL = "   Select A.ID,A.分类ID,A.编号,A.名称 From 病历词句分类 B, 病历词句示范 A" & _
        "   Where A.分类id = B.ID And Nvl(Substr(B.范围, [1], 1), '0') = '1' And " & strMatch & _
        "   And ((Nvl(A.通用级, 0) = 0" & _
        "       Or A.通用级 = 1 And A.科室id In (Select A.部门id From 部门人员 A, 上机人员表 B Where A.人员id = B.人员id And B.用户名 = User)" & _
        "       Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User)))" & strFind & _
        "   Order by " & IIF(lngTypeID = 0, "", " DECODE(A.分类ID," & lngTypeID & ",0,1),") & "A.分类ID,A.编号"
    Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
        CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")

    Call LocaItem
        
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub LocaItem()
    Dim lngRow As Long
    
    If mrsFind.RecordCount = 0 Then
        lblInfo.Caption = "没有找到符合条件的信息"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    
    If mrsFind.EOF = True Then
        lblInfo.Caption = "已经完成所有定位，请重新输入条件"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    lblInfo.Caption = "共找到" & mrsFind.RecordCount & "条,当前是第" & mrsFind.AbsolutePosition & "条"
    lblInfo.ForeColor = &H8000000D
    
    If mrsFind.RecordCount > 0 Then
        If mrsFind.RecordCount <> mrsFind.AbsolutePosition Then
            cmdFind.Caption = "下一个(&L)"
        Else
            cmdFind.Caption = "定位(&L)"
            lblInfo.Caption = "已经是最后一条，请重新输入条件"
        End If
    End If
    
    '开始进行定位
    tvw_s.Nodes("_" & mrsFind!分类id).Selected = True
    tvw_s.SelectedItem.EnsureVisible
    Call ShowList(mrsFind!分类id)
    
    For lngRow = vsList.FixedRows To vsList.Rows - 1
        If Val(vsList.RowData(lngRow)) = Val(mrsFind!ID) Then
            vsList.ROW = lngRow
            vsList.TopRow = lngRow
            Exit For
        End If
    Next lngRow
End Sub

Private Sub cmdGO_Click()
    Dim blnTrans As Boolean
    Dim rsTemp As ADODB.Recordset
    If Not UpdateFormat Then
        txtAdvice.SetFocus: Exit Sub
    End If
    On Error GoTo ErrHand
    mrsFormat.Filter = 0
    gcnOracle.BeginTrans: blnTrans = True
    With mrsFormat
        Do While Not .EOF
            If Val(!是否修改 & "") = 1 Then
                gstrSQL = "Zl_护理内容导入定义_Update(" & !类别 & ",'" & !名称 & "','" & Replace(!新格式, "'", "''") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        .MoveNext
        Loop
    End With
    gcnOracle.CommitTrans: blnTrans = False
    '保存后，则更原有记录信息
    mrsFormat.Filter = 0
    Set rsTemp = zlDatabase.CopyNewRec(mrsFormat)
    rsTemp.Filter = 0
    Do While Not rsTemp.EOF
        If Val(rsTemp!是否修改 & "") = 1 Then
            mrsFormat.Filter = "类别=" & rsTemp!类别
            mrsFormat!格式 = rsTemp!新格式 & ""
            mrsFormat!是否修改 = 0
            mrsFormat!新格式 = ""
            mrsFormat.Update
        End If
        rsTemp.MoveNext
    Loop
    picDef.Visible = False
    SetControlEnable True
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    If rtfSentence.Text = "" Then
        MsgBox "没有可用的词句内容。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrSentence = Replace(Replace(rtfSentence.Text, "|", "∣"), "'", "")
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF3 Then
        If cmdFind.Enabled And cmdFind.Visible Then Call cmdFind_Click
    ElseIf KeyCode = vbKeyA And Shift = vbAltMask And picDef.Visible = True Then
        Call cmdGO_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim vRect As RECT, lngMaxH As Long
    Dim rsTemp As New ADODB.Recordset
    
    mblnShow = True
    mblnOK = False
    mstrSentence = ""
    Me.rtfSentence.Text = mstrInput
    
    On Error GoTo errH
    If mobjPublicLis Is Nothing Then
        On Error Resume Next
        Set mobjPublicLis = CreateObject("zlPublicLIS.clsSampleReprot")
        Err.Clear: On Error GoTo 0
        If Not mobjPublicLis Is Nothing Then
            Call mobjPublicLis.InitSampleReprot(gcnOracle, glngSys, 1265, "")
        End If
    End If
    If mobjPublicLis Is Nothing Then
        MsgBox "LIS公共部件zlPublicLIS创建失败，将不能查看导入检验结果！", vbInformation, gstrSysName
    End If
    '导入内容自定义设置
    '读取自定格式
    gstrSQL = "Select 类别,名称,格式 from 护理内容导入定义"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "护理内容导入定义")
    Set mrsFormat = zlDatabase.CopyNewRec(rsTemp, , , Array("是否修改", adInteger, 1, 0, "新格式", adVarChar, 500, Empty))
    
    txtAdvice.Tag = ""
    Call Record_Init(mrsField, "类别," & adInteger & ",1|字段," & adVarChar & ",2000")
    mrsField.AddNew: mrsField!类别 = 1: mrsField!字段 = "[编号],[名称],[内容文本]" '病例词句
    mrsField.AddNew: mrsField!类别 = 2: mrsField!字段 = "[开始时间],[开单医生],[医嘱内容],[诊疗项目],[单量],[总量],[医生嘱托],[给药途径]" '医嘱内容
    mrsField.AddNew: mrsField!类别 = 3: mrsField!字段 = "[指标代码],[指标中文名],[检验结果],[单位],[结果标志],[结果参考]" '检验项目(普通项目)
    mrsField.AddNew: mrsField!类别 = 4: mrsField!字段 = "[细菌名],[耐药机制],[抗生素],[抗生素结果],[耐药性],[药敏方法],[用法用量1],[用法用量2],[血药浓度1],[血药浓度2],[尿药浓度1],[尿药浓度2]" '检验项目(微生物项目)
    mrsField.UpdateBatch
    With cbo类别
        .Clear
        .AddItem "护理词句": .ItemData(.NewIndex) = 1
        .AddItem "医嘱内容": .ItemData(.NewIndex) = 2
        .AddItem "检验项目[普通项目]": .ItemData(.NewIndex) = 3
        .AddItem "检验项目[微生物项目]": .ItemData(.NewIndex) = 4
        .ListIndex = 0
    End With
    mintIndex = cbo类别.ListIndex
   
    
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gstrSQL = "Select B.主页ID as 就诊ID,NVL(B.性别,A.性别) 性别,Nvl(B.婚姻状况,A.婚姻状况) as 婚姻状况," & _
        " B.住院目的,B.当前病况 as 病人病情,B.入院方式" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlng病人ID, mlng主页ID)
    '读取词句数据
    Call ShowTree
    
    '界面显示处理
    Call RestoreWinState(Me, App.ProductName, IIF(mstrInput <> "", 1, 0))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    tvw_s.Left = 0
    tvw_s.Top = 0
    tvw_s.Height = Me.ScaleHeight - picBottom.Height
    
    fraLR.Left = tvw_s.Left + tvw_s.Width
    fraLR.Top = 0
    fraLR.Height = tvw_s.Height
    
    vsList.Top = 0
    vsList.Left = fraLR.Left + fraLR.Width
    vsList.Height = Me.ScaleHeight - rtfSentence.Height - fraUD.Height - picBottom.Height
    vsList.Width = Me.ScaleWidth - fraLR.Width - tvw_s.Width
    
    fraUD.Top = vsList.Top + vsList.Height
    fraUD.Left = vsList.Left
    fraUD.Width = vsList.Width
    
    rtfSentence.Top = fraUD.Top + fraUD.Height
    rtfSentence.Left = vsList.Left
    rtfSentence.Width = vsList.Width
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    If Not mrsPati Is Nothing Then
        If mrsPati.State = adStateOpen Then mrsPati.Close
        Set mrsPati = Nothing
    End If
    If Not mrsFind Is Nothing Then
        If mrsFind.State = adStateOpen Then mrsFind.Close
        Set mrsFind = Nothing
    End If
    If Not mrsField Is Nothing Then
        If mrsField.State = adStateOpen Then mrsField.Close
        Set mrsField = Nothing
    End If
    If Not mrsFormat Is Nothing Then
        If mrsFormat.State = adStateOpen Then mrsFormat.Close
        Set mrsFormat = Nothing
    End If
    Set mobjPublicLis = Nothing
    Set mobjXML = Nothing
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Call SaveWinState(Me, App.ProductName, IIF(mstrInput <> "", 1, 0))
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 1 Then
            If Me.Width + X < 4000 Or Me.Width + X > 9600 Then Exit Sub
            Me.Width = Me.Width + X
        ElseIf Index = 2 Then
            If Me.Height + Y < rtfSentence.Height * 2 Or Me.Height + Y > 7200 Then Exit Sub
            Me.Height = Me.Height + Y
        End If
        Call Form_Resize
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsList.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        
        vsList.Left = vsList.Left + X
        vsList.Width = vsList.Width - X
        
        fraUD.Left = fraUD.Left + X
        fraUD.Width = fraUD.Width - X
        
        rtfSentence.Left = rtfSentence.Left + X
        rtfSentence.Width = rtfSentence.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub fraUD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreY = Y
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If vsList.Height + (Y - mlngPreY) < 1000 Or rtfSentence.Height - (Y - mlngPreY) < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + (Y - mlngPreY)
        vsList.Height = vsList.Height + (Y - mlngPreY)
        rtfSentence.Top = rtfSentence.Top + (Y - mlngPreY)
        rtfSentence.Height = rtfSentence.Height - (Y - mlngPreY)
        
        Me.Refresh
    End If
End Sub

Private Sub imgClose_Click()
    Call cmdClose_Click
End Sub

Private Sub imgTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picTip_MouseMove(Button, Shift, X, Y)
End Sub
    
Private Sub picBottom_GotFocus()
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    
    If picBottom.ScaleWidth - cmdCancel.Width * 2 < 3500 Then Exit Sub
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
End Sub

Private Sub picTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    strInfo = "缺省格式内容" & vbCrLf & "  护理词句：[内容文本]" & vbCrLf & _
        "  医嘱内容：[医嘱内容]+[单量]+[给药途径]" & vbCrLf & _
        "  检验项目[普通项目]：[指标中文名]+""(""+[检验结果]+"")""" & vbCrLf & _
        "  检验项目[微生物项目]：[抗生素]+""(""+[抗生素结果]+"")"""
    Call zlCommFun.ShowTipInfo(picTip.hWnd, strInfo, True, True)
End Sub


Private Sub rtfSentence_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub tvw_s_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Children = 1 Then
        Node.Child.Expanded = True
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Val(Mid(Node.Key, 2)) <> 0 Then
        Call ShowList(Val(Mid(Node.Key, 2)))
    Else
        vsList.Rows = vsList.FixedRows
    End If
End Sub

Private Sub txtAdvice_Change()
    txtAdvice.Tag = "1"
End Sub


Private Sub txtFind_Change()
    If Trim(txtFind.Text) = "" Then
        lblInfo.Caption = "请输入查找条件"
        lblInfo.ForeColor = &H8000&
    Else
        lblInfo.Caption = "点击定位完成词句查找"
        lblInfo.ForeColor = &H8000000D
    End If
    
    cmdFind.Caption = "定位(&L)"
    Set mrsFind = New ADODB.Recordset
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdFind.SetFocus
        Call cmdFind_Click
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsList_DblClick()
    With vsList
        If .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call LoadWords
        End If
    End With
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call vsList_DblClick
    End If
End Sub

Private Sub LoadWords()
    Dim lngStart As Long, lngStart_LAST As Long
    Dim strText As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsValue As New ADODB.Recordset
    Dim bln是否微生物 As Boolean, arrTag() As String
    Dim strReturn As String
    On Error GoTo ErrHand
    
    If Val(vsList.RowData(vsList.ROW)) = 0 Then Exit Sub
    lngStart_LAST = rtfSentence.SelStart
    If lngStart_LAST = 0 Then lngStart_LAST = Len(rtfSentence.Text)
    rtfSentence.Tag = rtfSentence.Text
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.addObject "clsScript", mobjScript, True
        End If
    End If
    
    If Mid(tvw_s.SelectedItem.Key, 1, 1) = "_" Then
        gstrSQL = "Select 内容性质,内容文本,要素名称,要素单位 From 病历词句组成 Where 词句ID=[1] Order by 排列次序"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, Val(vsList.RowData(vsList.ROW)))
        
        rtfSentence.Text = ""
        Do While Not rsTemp.EOF
            lngStart = Len(rtfSentence.Text)
            rtfSentence.SelStart = lngStart
            rtfSentence.SelLength = 0
            Select Case rsTemp!内容性质
            Case 0 '自由文字
                strText = NVL(rsTemp!内容文本)
                With rtfSentence
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = False
                End With
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                If Not IsNull(rsTemp!内容文本) Then
                    strText = rsTemp!内容文本
                Else
                    strText = ""
                    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4]) as 内容 From Dual"
                    Set rsValue = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(rsTemp!要素名称), mlng病人ID, mlng主页ID, 2)
                    If Not rsTemp.EOF Then strText = IIF(Not IsNull(rsValue!内容), rsValue!内容 & NVL(rsTemp!要素单位), "")
                    If strText = "" Then strText = "{" & rsTemp!要素名称 & "}" & NVL(rsTemp!要素单位)
                End If
                With rtfSentence
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            rsTemp.MoveNext
        Loop
        strReturn = ""
        mrsFormat.Filter = "类别=1"
        If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!格式 & ""
        If strReturn <> "" Then
            If InStr(strReturn, "[编号]") > 0 Then
               strReturn = Replace(strReturn, "[编号]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("编号")) & """")
            End If
            If InStr(strReturn, "[名称]") > 0 Then
               strReturn = Replace(strReturn, "[名称]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("名称")) & """")
            End If
            If InStr(strReturn, "[内容文本]") > 0 Then
               strReturn = Replace(strReturn, "[内容文本]", """" & rtfSentence.Text & """")
            End If
            strReturn = mobjVBA.Eval(strReturn)
            rtfSentence.Text = strReturn
        End If
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "=" Then '医嘱
        strReturn = ""
        mrsFormat.Filter = "类别=2"
        If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!格式 & ""
        If strReturn = "" Then
            strReturn = vsList.TextMatrix(vsList.ROW, vsList.ColIndex("内容文本"))
        Else
            '"[开始时间],[开单医生],[医嘱内容],[诊疗项目],[单量],[总量],[医生嘱托],[给药途径]"
            If InStr(strReturn, "[开始时间]") > 0 Then
               strReturn = Replace(strReturn, "[开始时间]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("开始时间")) & """")
            End If
            If InStr(strReturn, "[开单医生]") > 0 Then
               strReturn = Replace(strReturn, "[开单医生]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("开嘱医生")) & """")
            End If
            If InStr(strReturn, "[医嘱内容]") > 0 Then
               strReturn = Replace(strReturn, "[医嘱内容]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("医嘱内容")) & """")
            End If
            If InStr(strReturn, "[诊疗项目]") > 0 Then
               strReturn = Replace(strReturn, "[诊疗项目]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("诊疗项目")) & """")
            End If
            If InStr(strReturn, "[单量]") > 0 Then
               strReturn = Replace(strReturn, "[单量]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("单次用量")) & """")
            End If
            If InStr(strReturn, "[总量]") > 0 Then
               strReturn = Replace(strReturn, "[总量]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("总给予量")) & """")
            End If
            If InStr(strReturn, "[医生嘱托]") > 0 Then
               strReturn = Replace(strReturn, "[医生嘱托]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("医生嘱托")) & """")
            End If
            If InStr(strReturn, "[给药途径]") > 0 Then
               strReturn = Replace(strReturn, "[给药途径]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("给药途径")) & """")
            End If
            strReturn = mobjVBA.Eval(strReturn)
        End If
        rtfSentence.Text = strReturn
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "％" Then
        arrTag = Split(tvw_s.SelectedItem.Tag, "'")
        bln是否微生物 = Val(arrTag(4)) = 1
        strReturn = ""
        If bln是否微生物 = False Then
            mrsFormat.Filter = "类别=3"
            If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!格式 & ""
            If strReturn = "" Then '默认导入指标的名称和结果
                strReturn = vsList.TextMatrix(vsList.ROW, vsList.ColIndex("指标中文名"))
                If vsList.TextMatrix(vsList.ROW, vsList.ColIndex("指标中文名")) <> "" Then
                    strReturn = strReturn & ":" & "(" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("检验结果")) & ")"
                End If
            Else
                '"[指标代码],[指标中文名],[检验结果],[结果标志],[结果参考]"
                 If InStr(strReturn, "[指标代码]") > 0 Then
                    strReturn = Replace(strReturn, "[指标代码]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("指标代码")) & """")
                 End If
                 If InStr(strReturn, "[指标中文名]") > 0 Then
                    strReturn = Replace(strReturn, "[指标中文名]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("指标中文名")) & """")
                 End If
                 If InStr(strReturn, "[检验结果]") > 0 Then
                    strReturn = Replace(strReturn, "[检验结果]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("检验结果")) & """")
                 End If
                 If InStr(strReturn, "[单位]") > 0 Then
                    strReturn = Replace(strReturn, "[单位]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("单位")) & """")
                 End If
                 If InStr(strReturn, "[结果标志]") > 0 Then
                    strReturn = Replace(strReturn, "[结果标志]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("结果标志")) & """")
                 End If
                 If InStr(strReturn, "[结果参考]") > 0 Then
                    strReturn = Replace(strReturn, "[结果参考]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("结果参考")) & """")
                 End If
                 strReturn = mobjVBA.Eval(strReturn)
            End If
            rtfSentence.Text = strReturn
        Else
            If vsList.RowOutlineLevel(vsList.ROW) <= 0 Then Exit Sub '微生物点击细菌列不处理，只能点击子项目
            mrsFormat.Filter = "类别=4"
            '"[细菌名],[耐药机制],[抗生素],[抗生素结果],[耐药性],[药敏方法],[用法用量1],[用法用量2],[血药浓度1],[血药浓度2],[尿药浓度1],[尿药浓度2]"
            If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!格式 & ""
            If strReturn = "" Then '默认导入指标的名称和结果
                strReturn = vsList.TextMatrix(vsList.ROW, vsList.ColIndex("抗生素"))
                If vsList.TextMatrix(vsList.ROW, vsList.ColIndex("抗生素结果")) <> "" Then
                    strReturn = strReturn & ":" & "(" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("抗生素结果")) & ")"
                End If
            Else
                If InStr(strReturn, "[细菌名]") > 0 Then
                   strReturn = Replace(strReturn, "[细菌名]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("细菌名")) & """")
                End If
                If InStr(strReturn, "[耐药机制]") > 0 Then
                   strReturn = Replace(strReturn, "[耐药机制]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("耐药机制")) & """")
                End If
                If InStr(strReturn, "[抗生素]") > 0 Then
                   strReturn = Replace(strReturn, "[抗生素]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("抗生素")) & """")
                End If
                If InStr(strReturn, "[抗生素结果]") > 0 Then
                   strReturn = Replace(strReturn, "[抗生素结果]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("抗生素结果")) & """")
                End If
                If InStr(strReturn, "[耐药性]") > 0 Then
                   strReturn = Replace(strReturn, "[耐药性]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("耐药性")) & """")
                End If
                If InStr(strReturn, "[药敏方法]") > 0 Then
                   strReturn = Replace(strReturn, "[药敏方法]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("药敏方法")) & """")
                End If
                If InStr(strReturn, "[用法用量1]") > 0 Then
                   strReturn = Replace(strReturn, "[用法用量1]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("用法用量1")) & """")
                End If
                If InStr(strReturn, "[用法用量2]") > 0 Then
                   strReturn = Replace(strReturn, "[用法用量2]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("用法用量2")) & """")
                End If
                If InStr(strReturn, "[血药浓度1]") > 0 Then
                   strReturn = Replace(strReturn, "[血药浓度1]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("血药浓度1")) & """")
                End If
                If InStr(strReturn, "[血药浓度2]") > 0 Then
                   strReturn = Replace(strReturn, "[血药浓度2]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("血药浓度2")) & """")
                End If
                If InStr(strReturn, "[尿药浓度1]") > 0 Then
                   strReturn = Replace(strReturn, "[尿药浓度1]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("尿药浓度1")) & """")
                End If
                If InStr(strReturn, "[尿药浓度2]") > 0 Then
                  strReturn = Replace(strReturn, "[尿药浓度2]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("尿药浓度2")) & """")
                End If
                strReturn = mobjVBA.Eval(strReturn)
            End If
            rtfSentence.Text = strReturn
        End If
    End If
    
    rtfSentence.Text = Mid(rtfSentence.Tag, 1, lngStart_LAST) & "，" & rtfSentence.Text & Mid(rtfSentence.Tag, lngStart_LAST + 1) & "，"
    If Mid(rtfSentence.Text, 1, 1) = "，" Then rtfSentence.Text = Mid(rtfSentence.Text, 2)
    If Right(rtfSentence.Text, 1) = "，" Then rtfSentence.Text = Mid(rtfSentence.Text, 1, Len(rtfSentence.Text) - 1)
    If lngStart_LAST = Len(rtfSentence.Tag) Then lngStart_LAST = Len(rtfSentence.Text)
    rtfSentence.SelStart = lngStart_LAST
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowLisList()
'功能：根据选择的检验项目，展示结果信息
    Dim lngKey As Long '检验报告id
    Dim strXMLLIS As String '返回的结果信息
    Dim strTag As String, arrTag() As String
    Dim L_Item As LisItem
    Dim objXMLNodeList As Object, objXMLNode As Object, objChildNode As Object, objPChildNode As Object
    Dim strFirstName As String, strChildFirstName As String
    Dim lngStartRow As Long
    Dim strTmp As String, i As Integer
    '120692:添加检验项目
    If mobjPublicLis Is Nothing Then Exit Sub
    If tvw_s.SelectedItem.Key = "％" Then '父节点
        '按照检验项目和指标分布
    Else
        lngKey = Val(Mid(tvw_s.SelectedItem.Key, 2))
'        objNode.Tag = L_Item.检验报告id & "'" & L_Item.申请id & "'" & L_Item.紧急标志 & "'" & L_Item.标本序号 & "'" & L_Item.是否微生物 & "'" & _
'                        L_Item.检验次数 & "'" & L_Item.检验人 & "'" & L_Item.审核人 & "'" & L_Item.审核时间 & "'" & L_Item.申请时间
        strTag = tvw_s.SelectedItem.Tag
        arrTag = Split(strTag, "'")
        L_Item.是否微生物 = arrTag(4)
        Call InitVsf(2, L_Item.是否微生物 = 1)
        strXMLLIS = mobjPublicLis.GetLaboratoryReportResultList(lngKey)
        If strXMLLIS <> "" Then
            If OpenXMLDocument(strXMLLIS) = True Then
                If L_Item.是否微生物 = 0 Then
                    Set objXMLNodeList = mobjXML.selectNodes(".//普通项目//指标内容").Item(0).childNodes
                    strFirstName = objXMLNodeList.Item(0).nodename
                    vsList.Redraw = flexRDNone
                    For Each objXMLNode In objXMLNodeList
                        If objXMLNode.nodename = strFirstName Then
                            vsList.Rows = vsList.Rows + 1
                            vsList.Cell(flexcpPicture, vsList.Rows - 1, 0) = imgList.ListImages(2).Picture
                        End If
                        Select Case objXMLNode.nodename
                            Case "指标id"
                                vsList.RowData(vsList.Rows - 1) = objXMLNode.Text
                            Case "指标代码", "指标英文名", "指标中文名", "检验结果", "单位", "结果标志", "结果参考"
                                vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex(objXMLNode.nodename)) = objXMLNode.Text
                        End Select
                    Next
                    vsList.Redraw = flexRDDirect
                Else '微生物项目
                    Set objXMLNodeList = mobjXML.selectNodes(".//微生物项目").Item(0).childNodes
                    strFirstName = objXMLNodeList.Item(0).nodename
                    vsList.Redraw = flexRDNone
                    For Each objXMLNode In objXMLNodeList
                        If objXMLNode.nodename = strFirstName Then
                            vsList.Rows = vsList.Rows + 1
                            vsList.MergeRow(vsList.Rows - 1) = True
                            vsList.Cell(flexcpPicture, vsList.Rows - 1, 0) = imgList.ListImages(2).Picture
                            lngStartRow = vsList.Rows - 1
                            strTmp = ""
                        End If
                        Select Case objXMLNode.nodename
                            Case "细菌id"
                                vsList.RowData(vsList.Rows - 1) = objXMLNode.Text
                            Case "细菌名", "描述", "耐药机制"
                                vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex(objXMLNode.nodename)) = objXMLNode.Text
                            Case "抗生素结果列表"
                                vsList.IsSubtotal(lngStartRow) = True
                                vsList.RowOutlineLevel(lngStartRow) = 0
                                strTmp = vsList.TextMatrix(lngStartRow, vsList.ColIndex("细菌名")) & "[" & vsList.TextMatrix(lngStartRow, vsList.ColIndex("描述")) & "]"
                                vsList.TextMatrix(lngStartRow, vsList.ColIndex("细菌名称")) = strTmp
                                '具体的抗生素和结果
                                Set objPChildNode = objXMLNode.childNodes
                                strChildFirstName = objPChildNode.Item(0).nodename
                                For Each objChildNode In objPChildNode
                                    If objChildNode.nodename = strChildFirstName Then
                                        vsList.Rows = vsList.Rows + 1
                                        vsList.MergeRow(vsList.Rows - 1) = False
                                        vsList.RowData(vsList.Rows - 1) = vsList.RowData(lngStartRow)
                                        vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex("细菌名")) = vsList.TextMatrix(lngStartRow, vsList.ColIndex("细菌名"))
                                        vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex("描述")) = vsList.TextMatrix(lngStartRow, vsList.ColIndex("描述"))
                                        vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex("耐药机制")) = vsList.TextMatrix(lngStartRow, vsList.ColIndex("耐药机制"))
                                        vsList.Cell(flexcpPicture, vsList.Rows - 1, 0) = imgList.ListImages(2).Picture
                                        vsList.IsSubtotal(vsList.Rows - 1) = True
                                        vsList.RowOutlineLevel(vsList.Rows - 1) = 1
                                        vsList.IsCollapsed(vsList.Rows - 1) = flexOutlineExpanded
                                    End If
                                    Select Case objChildNode.nodename
                                        Case "抗生素", "抗生素结果", "耐药性", "药敏方法", "用法用量1", "用法用量2", "血药浓度1", "血药浓度2", "尿药浓度1", "尿药浓度2"
                                            vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex(objChildNode.nodename)) = objChildNode.Text
                                            vsList.TextMatrix(lngStartRow, vsList.ColIndex(objChildNode.nodename)) = strTmp
                                    End Select
                                Next
                        End Select
                    Next
                    For i = vsList.ColIndex("细菌名称") To vsList.Cols - 1
                        vsList.MergeCol(i) = True
                    Next
                    vsList.Redraw = flexRDDirect
                End If
            End If
        End If
    End If
End Sub

Private Sub InitVsf(ByVal intType As Integer, Optional bln是否微生物 As Boolean = False)
'功能：初始化表格信息
'intType:0-词句选择,1-医嘱,2-检验 (intType-2时需要传入是否微生物)
    With vsList
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 1
        .Cols = 0
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarSimpleLeaf
        .Editable = flexEDNone
        .MergeCells = flexMergeNever
        Select Case intType
            Case 0
                .Cols = 4
                .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                .TextMatrix(0, 1) = "编号": .ColKey(1) = "编号": .ColWidth(1) = 795
                .TextMatrix(0, 2) = "名称": .ColKey(2) = "名称": .ColWidth(2) = 1530
                .TextMatrix(0, 3) = "内容": .ColKey(3) = "内容": .ColWidth(3) = 2535
            Case 1
                .Cols = 12
                .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                .TextMatrix(0, 1) = "相关ID": .ColKey(1) = "相关ID": .ColWidth(1) = 0: .ColHidden(1) = True
                .TextMatrix(0, 2) = "": .ColKey(2) = "一组": .ColWidth(2) = 315
                .TextMatrix(0, 3) = "医嘱内容": .ColKey(3) = "医嘱内容": .ColWidth(3) = 4000
                .TextMatrix(0, 4) = "单次用量": .ColKey(4) = "单次用量": .ColWidth(4) = 900
                .TextMatrix(0, 5) = "给药途径": .ColKey(5) = "给药途径": .ColWidth(5) = 1400
                .TextMatrix(0, 6) = "总给予量": .ColKey(6) = "总给予量": .ColWidth(6) = 795
                .TextMatrix(0, 7) = "诊疗项目": .ColKey(7) = "诊疗项目": .ColWidth(7) = 2000
                .TextMatrix(0, 8) = "开嘱医生": .ColKey(8) = "开嘱医生": .ColWidth(8) = 900
                .TextMatrix(0, 9) = "开始时间": .ColKey(9) = "开始时间": .ColWidth(9) = 1600
                .TextMatrix(0, 10) = "医生嘱托": .ColKey(10) = "医生嘱托": .ColWidth(10) = 1500
                .TextMatrix(0, 11) = "内容文本": .ColKey(11) = "内容文本": .ColWidth(11) = 0: .ColHidden(11) = True 'a.医嘱内容 || a.单次用量 || b.计算单位 || c.医嘱内容
            Case 2
                If bln是否微生物 = False Then
                    .Cols = 8
                    .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                    .TextMatrix(0, 1) = "指标代码": .ColKey(1) = "指标代码": .ColWidth(1) = 900
                    .TextMatrix(0, 2) = "指标英文名": .ColKey(2) = "指标英文名": .ColWidth(2) = 1530: .ColHidden(2) = True
                    .TextMatrix(0, 3) = "指标中文名": .ColKey(3) = "指标中文名": .ColWidth(3) = 3000
                    .TextMatrix(0, 4) = "检验结果": .ColKey(4) = "检验结果": .ColWidth(4) = 1200
                    .TextMatrix(0, 5) = "单位": .ColKey(5) = "单位": .ColWidth(5) = 900
                    .TextMatrix(0, 6) = "结果标志": .ColKey(6) = "结果标志": .ColWidth(6) = 900
                    .TextMatrix(0, 7) = "结果参考": .ColKey(7) = "结果参考": .ColWidth(7) = 900
                Else
                    .Cols = 15
                    .OutlineCol = 4
                    .MergeCells = flexMergeRestrictRows
                    .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                    .TextMatrix(0, 1) = "描述": .ColKey(1) = "描述": .ColWidth(1) = 0: .ColHidden(1) = True
                    .TextMatrix(0, 2) = "耐药机制": .ColKey(2) = "耐药机制": .ColWidth(2) = 0: .ColHidden(2) = True
                    .TextMatrix(0, 3) = "细菌名": .ColKey(3) = "细菌名": .ColWidth(3) = 0: .ColHidden(3) = True
                    .TextMatrix(0, 4) = "细菌名": .ColKey(4) = "细菌名称": .ColWidth(4) = 900
                    .TextMatrix(0, 5) = "抗生素": .ColKey(5) = "抗生素": .ColWidth(5) = 3000
                    .TextMatrix(0, 6) = "抗生素结果": .ColKey(6) = "抗生素结果": .ColWidth(6) = 1500
                    .TextMatrix(0, 7) = "耐药性": .ColKey(7) = "耐药性": .ColWidth(7) = 900
                    .TextMatrix(0, 8) = "药敏方法": .ColKey(8) = "药敏方法": .ColWidth(8) = 1200
                    .TextMatrix(0, 9) = "用法用量1": .ColKey(9) = "用法用量1": .ColWidth(9) = 1200
                    .TextMatrix(0, 10) = "用法用量2": .ColKey(10) = "用法用量2": .ColWidth(10) = 1200
                    .TextMatrix(0, 11) = "血药浓度1": .ColKey(11) = "血药浓度1": .ColWidth(11) = 1200
                    .TextMatrix(0, 12) = "血药浓度2": .ColKey(12) = "血药浓度2": .ColWidth(12) = 1200
                    .TextMatrix(0, 13) = "尿药浓度1": .ColKey(13) = "尿药浓度1": .ColWidth(13) = 1200
                    .TextMatrix(0, 14) = "尿药浓度2": .ColKey(14) = "尿药浓度2": .ColWidth(14) = 1200
                End If
        End Select
    End With
End Sub

Private Function OpenXMLDocument(ByVal strXml As String) As Boolean
    '******************************************************************************************************************
    '功能：打开XML文档
    '参数：strXML-XML字符串
    '返回：成功返回True，否则返回False
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    mstrXmlVersion = GetXMLVersion
    
    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    OpenXMLDocument = mobjXML.loadXML(strXml)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    Set mobjXML = Nothing
    OpenXMLDocument = False
End Function

Private Function GetXMLVersion() As String
    
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo ErrHand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
    If OS.IsDesinMode = True Or zlRegInfo("授权性质") <> "1" Then
        For intLoop = 0 To UBound(varXMLVersion)
            Err = 0
            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If Err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        On Error GoTo ErrHand
        
        If strXMLVer = "" Then
            MsgBox "创建MSXML2.DOMDocument对象失败"
            Exit Function
        End If
    Else
        strXMLVer = ""
    End If
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    MsgBox Err.Description
End Function

Private Sub SetControlEnable(ByVal blnEnable As Boolean)
'功能：打开自定义设置界面，则设置出该界面上的其他控件不可用，取消则恢复
'blnEnable =false 表示是显示自定义界面,True表示关闭自定义界面
    Dim objControl As Object
    For Each objControl In Me.Controls
        If InStr(1, ",ImageList,Line,", "," & TypeName(objControl) & ",") = 0 Then
            If objControl.Visible = True Then
                '排除自定义列本身
                If InStr(1, ",picDef,picTip,imgTip,lblDefTitle,imgClose,lblPrompt,lbl类别,cbo类别,lbl内容格式,txtAdvice,lbl字段项目,cbo字段,cmdAdd,cmdCheck,cmdGO,cmdClose,", "," & objControl.Name & ",") = 0 Then
                    objControl.Enabled = blnEnable
                End If
            End If
        End If
    Next
End Sub
Private Sub SetTag一并给药(Optional ByVal lngRow As Long)
'功能：在一并给药的医嘱前加标志
    Dim i As Long
    Dim lngBg As Long, lngEd As Long
    Dim j As Long
    Dim lngStart As Long, lngEnd As Long
    With vsList
        If lngRow = 0 Then
            lngStart = .FixedRows
            lngEnd = .Rows - 1
        Else
            lngStart = lngRow
            lngEnd = lngRow
        End If
        For i = lngStart To lngEnd
             lngBg = -1: lngEd = -1
             If RowIn一并给药(i, lngBg, lngEd) Then
                For j = lngBg To lngEd
                    If j = lngBg Then
                        .TextMatrix(j, .ColIndex("一组")) = "┏"
                    ElseIf j = lngEd Then
                        .TextMatrix(j, .ColIndex("一组")) = "┗"
                    Else
                        .TextMatrix(j, .ColIndex("一组")) = "┃"
                    End If
                Next
                If lngEd <> -1 Then
                   i = lngEd + 1
                End If
            End If
        Next
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明:PASS 中的 “RowIn一并给药” 与此方法相同,修改此方法也需要同步修改 PASS同名方法
    Dim i As Long, blnTmp As Boolean
    With vsList
        If Val(.TextMatrix(lngRow - 1, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function



