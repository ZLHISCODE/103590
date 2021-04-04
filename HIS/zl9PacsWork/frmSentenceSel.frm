VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmSentenceSel 
   AutoRedraw      =   -1  'True
   Caption         =   "词句选择"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9360
   Icon            =   "frmSentenceSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   3465
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      ScaleWidth      =   9360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6030
      Width           =   9360
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5865
         TabIndex        =   7
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4770
         TabIndex        =   6
         Top             =   135
         Width           =   1100
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "详细内容"
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   30
         Width           =   720
      End
   End
   Begin RichTextLib.RichTextBox rtfSentence 
      Height          =   825
      Left            =   3555
      TabIndex        =   2
      Top             =   4680
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   1455
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmSentenceSel.frx":058A
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
      Width           =   5760
      _cx             =   10160
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
      FormatString    =   $"frmSentenceSel.frx":0627
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
               Picture         =   "frmSentenceSel.frx":069C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSentenceSel.frx":0C36
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSentenceSel.frx":11D0
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
            Picture         =   "frmSentenceSel.frx":176A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceSel.frx":1D04
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
Attribute VB_Name = "frmSentenceSel"
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
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
'===============================================================================================
Public mblnShow As Boolean '该窗体是否正在显示

Private mint来源 As Integer '病人来源：1-门诊,2-住院
Private mlng病人ID As Long
Private mvar就诊ID As Variant '主页ID或挂号单号
Private mlng项目ID As Long '当前医嘱诊疗项目ID
Private mstr检查部位 As String '用","分隔的检查部位名称串
Private mstr检查方法 As String '用","分隔的检查方法名称串
Private mstrInput As String
Private mlngInputHwnd As Long

Private mrsPati As ADODB.Recordset
Private mrsItem As ADODB.Recordset

Private mstrSentence As String
Private mstrLike As String
Private mblnOK As Boolean

Private mlngPreY As Long
Private mobjEmrInterface As Object           '新版病历申请附项读取部件

Public Function ShowMe(frmParent As Object, ByVal int来源 As Integer, ByVal lng病人ID As Long, ByVal var就诊ID As Variant, _
    ByVal lng项目ID As Long, ByVal str检查部位 As String, ByVal str检查方法 As String, _
    Optional ByVal strInput As String, Optional ByVal lngInputHwnd As Long, Optional blnCancel As Boolean, Optional objEmrInterface As Object) As String
    
    mint来源 = int来源
    mlng病人ID = lng病人ID
    mvar就诊ID = var就诊ID
    mlng项目ID = lng项目ID
    mstr检查部位 = str检查部位
    mstr检查方法 = str检查方法
    
    mstrInput = strInput
    mlngInputHwnd = lngInputHwnd
    Set mobjEmrInterface = objEmrInterface
    
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    
    If mblnOK Then
        ShowMe = mstrSentence
    Else
        blnCancel = True
    End If
End Function

Private Function ShowTree() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strMatch As String
    
    Dim str诊疗类别 As String, str检查类型 As String, strDeptIDs As String, strIF As String
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
        
    If Not mrsItem.EOF Then str诊疗类别 = mrsItem!诊疗类别: str检查类型 = mrsItem!检查类型
    
    strDeptIDs = "," & GetUser科室IDs & ","
    strIF = " And (Nvl(A.通用级, 0) = 0 Or A.通用级 = 1 And Instr([11],','||A.科室id||',')>0" & _
            " Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User))"
            
    strMatch = "f_Sentence_Matched(A.ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1"
    strSQL = _
        " Select Max(Level) As 级数, ID, 上级id, 编码, 名称, 说明" & _
        " From 病历词句分类" & _
        " Start With ID In (" & _
        "   Select A.分类id From 病历词句分类 B, 病历词句示范 A" & _
        "   Where A.分类id = B.ID And Nvl(Substr(B.范围, 8, 1), '0') = '1' And " & strMatch & strIF & _
        "   Group By A.分类id)" & _
        " Connect By Prior 上级id = ID" & _
        " Group By ID, 上级id, 编码, 名称, 说明" & _
        " Order By 级数 Desc, 编码"
    If Not mrsPati Is Nothing Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mint来源, CStr(Nvl(mrsPati!性别)), CStr(Nvl(mrsPati!婚姻状况)), _
            CStr(Nvl(mrsPati!住院目的)), CStr(Nvl(mrsPati!病人病情)), CStr(Nvl(mrsPati!入院方式)), str诊疗类别, str检查类型, mstr检查部位, mstr检查方法, strDeptIDs)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mint来源, "", "", "", "", "", str诊疗类别, str检查类型, mstr检查部位, mstr检查方法, strDeptIDs)
    End If
    
    tvw_s.Nodes.Clear
    Set objNode = tvw_s.Nodes.Add(, , "_", "所有词句", "Close")
    objNode.ExpandedImage = "Expend"
    objNode.Expanded = True
    
    Do While Not rsTmp.EOF
        Set objNode = tvw_s.Nodes.Add("_" & Nvl(rsTmp!上级ID), tvwChild, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Close")
        objNode.ExpandedImage = "Expend"
        'objNode.Expanded = True
        
        rsTmp.MoveNext
    Loop

    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Selected = True
    End If
    If Not tvw_s.SelectedItem Is Nothing Then
        tvw_s.SelectedItem.EnsureVisible
    End If
    
    Screen.MousePointer = 0
    ShowTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowList(Optional ByVal lng分类ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strMatch As String, strIF As String, strDeptIDs As String, strInput As String
    
    Dim str诊疗类别 As String, str检查类型 As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If Not mrsItem.EOF Then str诊疗类别 = mrsItem!诊疗类别: str检查类型 = mrsItem!检查类型
    
    strMatch = "f_Sentence_Matched(A.ID,[2],[3],[4],[5],[6],[7],[8],[9],[10],[11])=1"
    
    strDeptIDs = "," & GetUser科室IDs & ","
    strIF = " And (Nvl(A.通用级, 0) = 0 Or A.通用级 = 1 And Instr([12],','||A.科室id||',')>0" & _
            " Or A.通用级 = 2 And A.人员id In (Select 人员id From 上机人员表 Where 用户名 = User))"
    
    If lng分类ID <> 0 Then
        '按树形读取数据
        strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Trim(B.内容文本) as 内容文本" & _
            " From 病历词句组成 B,病历词句示范 A" & _
            " Where A.ID=B.词句ID And B.排列次序=1 And A.分类ID=[1] And " & strMatch & strIF & _
            " Order by A.编号"
        If Not mrsPati Is Nothing Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng分类ID, mint来源, CStr(Nvl(mrsPati!性别)), CStr(Nvl(mrsPati!婚姻状况)), _
                CStr(Nvl(mrsPati!住院目的)), CStr(Nvl(mrsPati!病人病情)), CStr(Nvl(mrsPati!入院方式)), str诊疗类别, str检查类型, mstr检查部位, mstr检查方法, strDeptIDs)
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng分类ID, mint来源, "", "", "", "", "", str诊疗类别, str检查类型, mstr检查部位, mstr检查方法, strDeptIDs)
        End If
    Else
        '按输入读取数据
        If IsNumeric(mstrInput) Then
            If Len(mstrInput) = 1 Then
                strIF = strIF & " And A.编号 = [1]"
                strInput = mstrInput
            Else
                strIF = strIF & " And A.编号 Like [1]"
                strInput = mstrInput & "%"
            End If
        Else
            strIF = strIF & " And (A.名称 Like [1] Or B.内容文本 Like [1])"
            strInput = IIF(Len(mstrInput) > 2, mstrLike, "") & mstrInput & "%"
        End If
        strSQL = "Select A.ID,A.编号,A.名称,A.通用级,LPad(B.排列次序,3,'0')||Trim(B.内容文本) as 内容文本" & _
            " From 病历词句分类 C,病历词句组成 B,病历词句示范 A" & _
            " Where A.ID=B.词句ID And Nvl(B.内容性质,0)=0 And A.分类ID=C.ID And Nvl(Substr(C.范围, 8, 1), '0') = '1'" & strIF
        
        strSQL = "Select A.ID,A.编号,A.名称,A.通用级,Substr(Min(A.内容文本),4) as 内容文本" & _
            " From (" & strSQL & ") A Where " & strMatch & " Group by A.ID,A.编号,A.名称,A.通用级 Order by A.编号"
        
        If Not mrsPati Is Nothing Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strInput, mint来源, CStr(Nvl(mrsPati!性别)), CStr(Nvl(mrsPati!婚姻状况)), _
                CStr(Nvl(mrsPati!住院目的)), CStr(Nvl(mrsPati!病人病情)), CStr(Nvl(mrsPati!入院方式)), str诊疗类别, str检查类型, mstr检查部位, mstr检查方法, strDeptIDs)
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strInput, mint来源, "", "", "", "", "", str诊疗类别, str检查类型, mstr检查部位, mstr检查方法, strDeptIDs)
        End If
    End If
        
    rtfSentence.Text = ""
    vsList.Redraw = flexRDNone
    vsList.Rows = vsList.FixedRows
    
    If Not rsTmp.EOF Then
        vsList.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            vsList.RowData(i) = Val(rsTmp!ID)
            vsList.TextMatrix(i, 1) = rsTmp!编号
            vsList.TextMatrix(i, 2) = rsTmp!名称
            vsList.TextMatrix(i, 3) = Nvl(rsTmp!内容文本)
            vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(Nvl(rsTmp!通用级, 0) + 1).Picture
            
            rsTmp.MoveNext
        Next
        vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
        vsList.Row = 1: vsList.Col = 2
    End If
    vsList.Redraw = flexRDDirect
    
    If vsList.Rows > vsList.FixedRows Then
        Call vsList_AfterRowColChange(-1, -1, vsList.Row, vsList.Col)
    End If
    
    Screen.MousePointer = 0
    ShowList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If rtfSentence.Text = "" Then
        MsgBox "没有可用的词句内容。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If rtfSentence.SelText = "" Then
        mstrSentence = rtfSentence.Text
    Else
        mstrSentence = rtfSentence.SelText
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim vRect As RECT, lngMaxH As Long
    
    On Error GoTo errH
    mblnShow = True
    mblnOK = False
    mstrSentence = ""
    
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    
    '读取项目信息
    strSQL = "Select B.名称 as 诊疗类别,A.操作类型 as 检查类型 From 诊疗项目目录 A,诊疗项目类别 B Where A.ID=[1] And A.类别=B.编码"
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng项目ID)
    
    '读取病人信息
    If mlng病人ID <> 0 And mvar就诊ID <> Empty Then
        If mint来源 = 1 Then
            strSQL = "Select B.ID as 就诊ID,Nvl(B.性别,A.性别) as 性别,A.婚姻状况," & _
                " Null as 住院目的,Null as 病人病情,Null as 入院方式" & _
                " From 病人信息 A,病人挂号记录 B " & _
                " Where A.病人ID=B.病人ID And b.记录状态 =1 and b.记录性质 =1 and A.病人ID=[1] And B.NO=[2]"
        Else
            strSQL = "Select B.主页ID as 就诊ID,A.性别,Nvl(B.婚姻状况,A.婚姻状况) as 婚姻状况," & _
                " B.住院目的,B.当前病况 as 病人病情,B.入院方式" & _
                " From 病人信息 A,病案主页 B" & _
                " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
        End If
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mvar就诊ID)
    End If
    
    '读取词句数据
    If mstrInput = "" Then
        Call ShowTree
    Else
        Call ShowList
    End If
    
    '界面显示处理
    Call RestoreWinState(Me, App.ProductName, IIF(mstrInput <> "", 1, 0))
    
    If mstrInput <> "" Then
        '无匹配数据或者唯一匹配时自动返回
        If vsList.Rows = vsList.FixedRows Then
            mblnOK = True: Unload Me: Exit Sub '非取消自动退出
        ElseIf vsList.Rows = vsList.FixedRows + 1 And vsList.Row = vsList.FixedRows _
            And vsList.RowData(vsList.Row) > 0 And rtfSentence.Text <> "" Then
            Call cmdOK_Click: Exit Sub '只有一个自动匹配退出
        End If
        
        '窗体样式设置
        Call zlControl.FormSetCaption(Me, False, False)
        tvw_s.Visible = False
        fraLR.Visible = False
        picBottom.Visible = False
        
        '边框设置
        For i = 0 To fraBorder.UBound
            fraBorder(i).BackColor = vbButtonFace
            fraBorder(i).Visible = True
            lin(i * 2).Visible = True
            lin(i * 2 + 1).Visible = True
        Next
        Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
        Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
        Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
        Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
        lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
        lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
        lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
        lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
        lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
        lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
        lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
        lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
        
        '窗体位置设置
        GetWindowRect mlngInputHwnd, vRect
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        lngMaxH = Screen.Height - vRect.Bottom - rtfSentence.Height - fraUD.Height - fraBorder(0).Height * 2 - 1000
        
        vsList.Height = vsList.Rows * vsList.RowHeightMin + 60
        If vsList.Height < 1000 Then vsList.Height = 1000
        If vsList.Height > lngMaxH Then vsList.Height = lngMaxH
        Me.Height = vsList.Height + rtfSentence.Height + fraUD.Height + fraBorder(0).Height * 2
        
        Me.Left = vRect.Left - fraBorder(0).Height
        Me.Top = vRect.Bottom
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    If mstrInput = "" Then
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
    ElseIf mstrInput <> "" Then
        fraBorder(0).Left = 0
        fraBorder(0).Top = 0
        fraBorder(0).Width = Me.ScaleWidth
        fraBorder(1).Top = fraBorder(0).Height
        fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
        fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
        fraBorder(2).Left = 0
        fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
        fraBorder(2).Width = Me.ScaleWidth
        fraBorder(3).Top = fraBorder(0).Height
        fraBorder(3).Left = 0
        fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
        
        vsList.Top = fraBorder(0).Height
        vsList.Left = fraBorder(0).Height
        vsList.Height = Me.ScaleHeight - rtfSentence.Height - fraUD.Height - fraBorder(0).Height * 2
        vsList.Width = Me.ScaleWidth - fraBorder(0).Height * 2
        
        fraUD.Top = vsList.Top + vsList.Height
        fraUD.Left = vsList.Left
        fraUD.Width = vsList.Width
        
        rtfSentence.Top = fraUD.Top + fraUD.Height
        rtfSentence.Left = vsList.Left
        rtfSentence.Width = vsList.Width
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    
    If Not mrsItem Is Nothing Then
        If mrsItem.State = 1 Then mrsItem.Close
    End If
    Set mrsItem = Nothing
    
    If Not mrsPati Is Nothing Then
        If mrsPati.State = 1 Then mrsPati.Close
    End If
    Set mrsPati = Nothing
    
    Call SaveWinState(Me, App.ProductName, IIF(mstrInput <> "", 1, 0))
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 1 Then
            If Me.Width + x < 4000 Or Me.Width + x > 9600 Then Exit Sub
            Me.Width = Me.Width + x
        ElseIf Index = 2 Then
            If Me.Height + Y < rtfSentence.Height * 2 Or Me.Height + Y > 7200 Then Exit Sub
            Me.Height = Me.Height + Y
        End If
        Call Form_Resize
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or vsList.Width - x < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + x
        tvw_s.Width = tvw_s.Width + x
        
        vsList.Left = vsList.Left + x
        vsList.Width = vsList.Width - x
        
        fraUD.Left = fraUD.Left + x
        fraUD.Width = fraUD.Width - x
        
        rtfSentence.Left = rtfSentence.Left + x
        rtfSentence.Width = rtfSentence.Width - x
        
        Me.Refresh
    End If
End Sub

Private Sub fraUD_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mlngPreY = Y
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width * 2
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
End Sub

Private Sub rtfSentence_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Val(Mid(Node.Key, 2)) <> 0 Then
        Call ShowList(Val(Mid(Node.Key, 2)))
    Else
        rtfSentence.Text = ""
        vsList.Rows = vsList.FixedRows
    End If
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset
    Dim rsValue As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngStart As Long, strText As String
    
    If NewRow = OldRow Or NewRow < vsList.FixedRows Then Exit Sub
    
    On Error GoTo errH
    
    strSQL = "Select 内容性质,内容文本,要素名称,要素单位 From 病历词句组成 Where 词句ID=[1] Order by 排列次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsList.RowData(vsList.Row)))
    
    rtfSentence.Text = ""
    
    Do While Not rsTmp.EOF
        lngStart = Len(rtfSentence.Text)
        rtfSentence.SelStart = lngStart
        rtfSentence.SelLength = 0
        Select Case rsTmp!内容性质
        Case 0 '自由文字
            strText = Nvl(rsTmp!内容文本)
            With rtfSentence
                .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                .SelUnderline = False
            End With
        Case 1, 2 '1-临时诊治要素,2-固定诊治要素
            If Not IsNull(rsTmp!内容文本) Then
                strText = rsTmp!内容文本
            Else
                strText = ""
                If mlng病人ID <> 0 And Not mrsPati Is Nothing Then
                    '如果是住院，则先取新版，没有值再取老版
                    If mint来源 = 2 Then
                        strText = GetOrderInspectInfo(mlng病人ID, CStr(rsTmp!要素名称))
                        If strText <> "" Then strText = strText & Nvl(rsTmp!要素单位)
                    End If
                    If strText = "" Then
                        strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4]) as 内容 From Dual"
                        Set rsValue = zlDatabase.OpenSQLRecord(strSQL, Me.Name, CStr(rsTmp!要素名称), mlng病人ID, Val(mrsPati!就诊ID), mint来源)
                        If Not rsValue.EOF Then strText = IIF(Not IsNull(rsValue!内容), rsValue!内容 & Nvl(rsTmp!要素单位), "")
                    End If
                End If
                If strText = "" Then strText = "{" & rsTmp!要素名称 & "}" & Nvl(rsTmp!要素单位)
            End If
            With rtfSentence
                .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                .SelUnderline = True
            End With
        End Select
        rsTmp.MoveNext
    Loop
    rtfSentence.SelStart = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsList_DblClick()
    With vsList
        If .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call cmdOK_Click
        End If
    End With
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Function GetOrderInspectInfo(ByVal lng病人ID As Long, ByVal strCondition As String) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        GetOrderInspectInfo = mobjEmrInterface.GetOrderInspectInfo(lng病人ID, strCondition)
    End If
    
End Function
