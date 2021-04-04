VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditAssistant 
   AutoRedraw      =   -1  'True
   Caption         =   "词句选择"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9825
   Icon            =   "frmEditAssistant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
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
      ScaleWidth      =   9825
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6030
      Width           =   9825
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
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   9
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6945
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
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmEditAssistant.frx":058A
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
      FormatString    =   $"frmEditAssistant.frx":0627
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
               Picture         =   "frmEditAssistant.frx":069C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":0C36
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":11D0
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
            Picture         =   "frmEditAssistant.frx":176A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditAssistant.frx":1D04
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
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(A.ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1"
    
    strSQL = _
        " Select 级数, Id, 上级id, 编码, 名称, 说明,分类id" & vbNewLine & _
        " From (With b As (Select a.分类id" & vbNewLine & _
        "                 From 病历词句分类 b, 病历词句示范 a" & vbNewLine & _
        "                 Where a.分类id = b.Id And Nvl(Substr(b.范围, [1], 1), '0') = '1' And " & strMatch & " And" & vbNewLine & _
        "                       ((Nvl(a.通用级, 0) = 0 Or" & vbNewLine & _
        "                       a.通用级 = 1 And" & vbNewLine & _
        "                       a.科室id In" & vbNewLine & _
        "                       (Select a.部门id From 部门人员 a, 上机人员表 b Where a.人员id = b.人员id And b.用户名 = User) Or" & vbNewLine & _
        "                       a.通用级 = 2 And a.人员id In (Select 人员id From 上机人员表 Where 用户名 = User)))" & vbNewLine & _
        "                 Group By a.分类id)" & vbNewLine & _
        "       Select Max(Level) As 级数, a.Id, a.上级id, a.编码, a.名称, a.说明, Max(b.分类id) 分类id" & vbNewLine & _
        "       From 病历词句分类 a, b" & vbNewLine & _
        "       Where a.Id = b.分类id(+)" & vbNewLine & _
        "       Start With a.Id In (b.分类id)" & vbNewLine & _
        "       Connect By Prior a.上级id = a.Id" & vbNewLine & _
        "       Group By a.Id, a.上级id, a.编码, a.名称, a.说明" & vbNewLine & _
        "       Order By 级数 Desc, 编码)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
        CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")
    
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
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=0", "其他类", "Close")
    objNode.ExpandedImage = "Expend"
    
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
        strMatch = "f_Sentence_Matched(A.ID,[2],[3],[4],[5],[6],[7],[8],[9],[10],[11])=1"
        If lng分类id <> 0 Then
            '按树形读取数据
            strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Trim(B.内容文本) as 内容文本" & _
                " From 病历词句组成 B,病历词句示范 A" & _
                " Where A.ID=B.词句ID(+) And B.排列次序(+)=1 And A.分类ID=[1] And " & strMatch & _
                "   And ((Nvl(A.通用级, 0) = 0" & _
                "       Or A.通用级 = 1 And A.科室id In(Select A.部门id From 部门人员 A, 上机人员表 B Where A.人员id = B.人员id And B.用户名 = User)" & _
                "       Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User))) Order by A.编号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng分类id, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
                CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")
        Else
            '按输入读取数据
            strSQL = "Select A.ID,A.编号,A.名称,A.通用级,LPad(B.排列次序,3,'0')||Trim(B.内容文本) as 内容文本" & _
                " From 病历词句分类 C,病历词句组成 B,病历词句示范 A" & _
                " Where A.ID=B.词句ID And Nvl(B.内容性质,0)=0 And A.分类ID=C.ID And Nvl(Substr(C.范围, [1], 1), '0') = '1'" & _
                "   And (A.编号 Like [1]||'%'" & _
                "       Or A.名称 Like " & IIf(mstrLike <> "", "'%'||", "") & "[1]||'%'" & _
                "       Or B.内容文本 Like " & IIf(mstrLike <> "", "'%'||", "") & "[1]||'%')" & _
                "   And ((Nvl(A.通用级, 0) = 0" & _
                "       Or A.通用级 = 1 And A.科室id In(Select A.部门id From 部门人员 A, 上机人员表 B Where A.人员id = B.人员id And B.用户名 = User)" & _
                "       Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User)))"
            
            strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Substr(Min(A.内容文本),4) as 内容文本" & _
                " From (" & strSQL & ") A Where " & strMatch & " Group by A.ID,A.编号,A.名称,A.通用级 Order by A.编号"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstrInput, mintType, CStr(NVL(mrsPati!性别)), CStr(NVL(mrsPati!婚姻状况)), _
                CStr(NVL(mrsPati!住院目的)), CStr(NVL(mrsPati!病人病情)), CStr(NVL(mrsPati!入院方式)), "", "", "", "")
        End If
    Else
        int执行分类 = lng分类id
        If tvw_s.SelectedItem.Key = "=" Then int执行分类 = 99
        strSQL = "" & _
            "SELECT ROWNUM AS ID,A.开嘱医生 AS 编号,A.执行时间方案 AS 名称,1 AS 通用级,A.医嘱内容||A.单次用量||B.计算单位||C.医嘱内容 AS 内容文本" & vbNewLine & _
            "FROM 病人医嘱记录 A,诊疗项目目录 B,病人医嘱记录 C,诊疗项目目录 D" & vbNewLine & _
            "WHERE A.诊疗类别 IN ('5','6','7') AND A.诊疗项目ID=B.ID AND A.病人ID=[1] AND A.主页ID=[2] And A.婴儿=[3]" & vbNewLine & _
            "AND C.诊疗类别='E' AND C.执行性质=1 AND D.ID=C.诊疗项目ID AND A.相关ID=C.ID AND NVL(D.执行分类,0)=[4]" & vbNewLine & _
            "AND C.上次执行时间 IS NOT NULL" & vbNewLine & _
            "ORDER BY A.开嘱时间 DESC"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID, mint婴儿, int执行分类)
    End If
        
    vsList.Redraw = flexRDNone
    vsList.Rows = vsList.FixedRows
    If Not rsTmp.EOF Then
        vsList.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            vsList.RowData(i) = Val(rsTmp!ID)
            vsList.TextMatrix(i, 1) = rsTmp!编号
            vsList.TextMatrix(i, 2) = rsTmp!名称
            vsList.TextMatrix(i, 3) = NVL(rsTmp!内容文本)
            vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(NVL(rsTmp!通用级, 0) + 1).Picture
            
            rsTmp.MoveNext
        Next
        vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
        vsList.ROW = 1: vsList.COL = 2
    End If
    vsList.Redraw = flexRDDirect
    
    Screen.MousePointer = 0
    ShowList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCanCel_Click()
    Unload Me
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
        "   Order by " & IIf(lngTypeID = 0, "", " DECODE(A.分类ID," & lngTypeID & ",0,1),") & "A.分类ID,A.编号"
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

Private Sub cmdOK_Click()
    If rtfSentence.Text = "" Then
        MsgBox "没有可用的词句内容。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrSentence = rtfSentence.Text
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF3 Then
        If cmdFind.Enabled And cmdFind.Visible Then Call cmdFind_Click
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim vRect As RECT, lngMaxH As Long
    
    mblnShow = True
    mblnOK = False
    mstrSentence = ""
    Me.rtfSentence.Text = mstrInput
    
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gstrSQL = "Select B.主页ID as 就诊ID,NVL(B.性别,A.性别) 性别,Nvl(B.婚姻状况,A.婚姻状况) as 婚姻状况," & _
        " B.住院目的,B.当前病况 as 病人病情,B.入院方式" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlng病人ID, mlng主页ID)
    '读取词句数据
    Call ShowTree
    
    '界面显示处理
    Call RestoreWinState(Me, App.ProductName, IIf(mstrInput <> "", 1, 0))
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
    Set mrsPati = Nothing
    Set mrsFind = Nothing
    Call SaveWinState(Me, App.ProductName, IIf(mstrInput <> "", 1, 0))
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

Private Sub picBottom_GotFocus()
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    
    If picBottom.ScaleWidth - cmdCancel.Width * 2 < 3500 Then Exit Sub
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
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
    On Error GoTo ErrHand
    
    lngStart_LAST = rtfSentence.SelStart
    If lngStart_LAST = 0 Then lngStart_LAST = Len(rtfSentence.Text)
    rtfSentence.Tag = rtfSentence.Text
    
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
                    If Not rsTemp.EOF Then strText = IIf(Not IsNull(rsValue!内容), rsValue!内容 & NVL(rsTemp!要素单位), "")
                    If strText = "" Then strText = "{" & rsTemp!要素名称 & "}" & NVL(rsTemp!要素单位)
                End If
                With rtfSentence
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            rsTemp.MoveNext
        Loop
    Else
        rtfSentence.Text = vsList.TextMatrix(vsList.ROW, 3)
    End If
    
    rtfSentence.Text = Mid(rtfSentence.Tag, 1, lngStart_LAST) & "，" & rtfSentence.Text & Mid(rtfSentence.Tag, lngStart_LAST + 1) & "，"
    If Mid(rtfSentence.Text, 1, 1) = "，" Then rtfSentence.Text = Mid(rtfSentence.Text, 2)
    If Right(rtfSentence.Text, 1) = "，" Then rtfSentence.Text = Mid(rtfSentence.Text, 1, Len(rtfSentence.Text) - 1)
    
    rtfSentence.SelStart = lngStart_LAST
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
