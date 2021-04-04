VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
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
      Enabled         =   -1  'True
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

Public mblnShow As Boolean '该窗体是否正在显示

Private mstrType As String  '病人主诉,家族史,现病史,体格全身检查,既往史
Private mlng病历文件id As Long
Private mstr性别 As String
Private mstr婚姻状况 As String

Private mstrInput As String
Private mlngInputHwnd As Long

Private mstrDepts As String '操作员所属的科室ID

Private mstrSentence As String
Private mblnOK As Boolean

Private mlngPreY As Long

Public Function ShowMe(frmParent As Object, ByVal lng病历文件id As Long, ByVal str性别 As String, ByVal str婚姻状况 As String, ByVal strType As String, _
    Optional ByVal strInput As String, Optional ByVal lngInputHwnd As Long, Optional blnCancel As Boolean) As String
    
    mlng病历文件id = lng病历文件id
    mstr性别 = str性别
    mstr婚姻状况 = str婚姻状况
    mstrType = strType
    
    mstrInput = strInput
    mlngInputHwnd = lngInputHwnd
    
    If mstrDepts = "" Then mstrDepts = GetUser科室IDs
    
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
    Dim strSql As String, i As Long
    Dim objNode As node, strMatch As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(A.ID,1,[1],[2],Null,Null,Null,Null)=1"
    strSql = _
        " Select Max(Level) As 级数, ID, 上级id, 编码, 名称, 说明" & _
        " From 病历词句分类" & _
        " Start With ID In (" & _
        "   Select A.分类id " & _
        "   From (Select Distinct A.词句分类id as id" & vbNewLine & _
        "       From 病历提纲词句 A, 病历文件结构 D, 病历文件结构 E" & vbNewLine & _
        "       Where E.文件id Is Null And E.对象类型 = 1 And E.内容文本 = [6] And E.ID = D.预制提纲id" & vbNewLine & _
        "             And D.文件id =[3] And D.对象类型 = 1 And D.ID = A.提纲id) B, 病历词句示范 A" & _
        "   Where " & strMatch & _
        "       And A.分类id = B.ID And (Nvl(A.通用级, 0) = 0 Or A.通用级 = 1 And Instr(','||[4]||',',','||A.科室id||',')>0 Or A.通用级 = 2 And A.人员id = [5])" & _
        "   Group By A.分类id)" & _
        " Connect By Prior 上级id = ID" & _
        " Group By ID, 上级id, 编码, 名称, 说明" & _
        " Order By 级数 Desc, 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, mstr性别, mstr婚姻状况, mlng病历文件id, mstrDepts, UserInfo.ID, mstrType)
    
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
    Dim strSql As String, i As Long
    Dim strMatch As String
    Dim strDept As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(A.ID,1,[1],[2],Null,Null,Null,Null)=1"
    
    If InStr(GetInsidePrivs(1070), "科室病历词句") <> 0 Then
        strDept = " And" & vbNewLine & _
                "      (Nvl(a.通用级, 0) = 0 Or" & vbNewLine & _
                "      a.通用级 In (1, 2) And" & vbNewLine & _
                "      a.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

     Else
        strDept = " And" & vbNewLine & _
                "      (Nvl(a.通用级, 0) = 0 Or" & vbNewLine & _
                "      a.通用级 = 1 And" & vbNewLine & _
                "      a.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      a.通用级 = 2 And a.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End If
    
    If lng分类ID <> 0 Then
        '按树形读取数据
        strSql = "Select A.ID, A.编号, A.名称, A.通用级, A.内容文本" & vbNewLine & _
                "From (Select A.ID, A.编号, A.名称, A.通用级, Trim(B.内容文本) As 内容文本" & vbNewLine & _
                "       From 病历词句示范 A, 病历词句组成 B," & vbNewLine & _
                "            (Select Distinct A.词句分类id" & vbNewLine & _
                "              From 病历提纲词句 A, 病历文件结构 D, 病历文件结构 E" & vbNewLine & _
                "              Where E.文件id Is Null And E.对象类型 = 1 And E.内容文本 = [6] And E.ID = D.预制提纲id And" & vbNewLine & _
                "                    D.文件id = [3] And D.对象类型 = 1 And D.ID = A.提纲id) D" & vbNewLine & _
                "       Where D.词句分类id = A.分类id And A.ID = B.词句id  And B.排列次序=1 And B.内容性质=0 And A.分类id=[7]" & vbNewLine & _
                strDept & vbNewLine & _
                "       ) A" & vbNewLine & _
                "Where " & strMatch & vbNewLine & _
                "Group By A.ID, A.编号, A.名称, A.通用级, A.内容文本" & vbNewLine & _
                "Order By A.编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, mstr性别, mstr婚姻状况, mlng病历文件id, mstrDepts, UserInfo.ID, mstrType, lng分类ID)
    Else
        '按输入读取数据
        strSql = "Select A.ID, A.编号, A.名称, A.通用级, Substr(Min(A.内容文本), 4) As 内容文本" & vbNewLine & _
                "From (Select A.ID, A.编号, A.名称, A.通用级, LPad(B.排列次序, 3, '0') || Trim(B.内容文本) As 内容文本" & vbNewLine & _
                "       From 病历词句示范 A, 病历词句组成 B, 病历词句分类 C," & vbNewLine & _
                "            (Select Distinct A.词句分类id" & vbNewLine & _
                "              From 病历提纲词句 A, 病历文件结构 D, 病历文件结构 E" & vbNewLine & _
                "              Where E.文件id Is Null And E.对象类型 = 1 And E.内容文本 = [6] And E.ID = D.预制提纲id And" & vbNewLine & _
                "                    D.文件id = [3] And D.对象类型 = 1 And D.ID = A.提纲id) D" & vbNewLine & _
                "       Where D.词句分类id = A.分类id And A.分类id = C.ID And A.ID = B.词句id And Nvl(B.内容性质, 0) = 0" & vbNewLine & _
                "       And (A.编号 Like [7] Or A.名称 Like [8] Or B.内容文本 Like [8])" & vbNewLine & _
                strDept & vbNewLine & _
                "       ) A" & vbNewLine & _
                "Where " & strMatch & vbNewLine & _
                "Group By A.ID, A.编号, A.名称, A.通用级" & vbNewLine & _
                "Order By A.编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, mstr性别, mstr婚姻状况, mlng病历文件id, mstrDepts, UserInfo.ID, _
                    mstrType, mstrInput & "%", gstrLike & mstrInput & "%")
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
    Dim strSql As String, i As Long
    Dim vRect As RECT, lngMaxH As Long
    
    mblnShow = True
    mblnOK = False
    mstrSentence = ""
        
    '读取词句数据
    If mstrInput = "" Then
        Call ShowTree
    Else
        Call ShowList
    End If
    
    '界面显示处理
    Call RestoreWinState(Me, App.ProductName, IIf(mstrInput <> "", 1, 0))
    
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
        lin(0).x1 = 0: lin(0).y1 = 0: lin(0).x2 = Screen.Width: lin(0).y2 = lin(0).y1: lin(0).BorderColor = &H8000000F
        lin(1).x1 = 0: lin(1).y1 = Screen.TwipsPerPixelY: lin(1).x2 = Screen.Width: lin(1).y2 = lin(1).y1: lin(1).BorderColor = &H8000000E
        lin(2).x1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).y1 = 0: lin(2).x2 = lin(2).x1: lin(2).y2 = Screen.Height: lin(2).BorderColor = &H80000011
        lin(3).x1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).y1 = 0: lin(3).x2 = lin(3).x1: lin(3).y2 = Screen.Height: lin(3).BorderColor = &H80000010
        lin(4).x1 = 0: lin(4).y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).x2 = Screen.Width: lin(4).y2 = lin(4).y1: lin(4).BorderColor = &H80000011
        lin(5).x1 = 0: lin(5).y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).x2 = Screen.Width: lin(5).y2 = lin(5).y1: lin(5).BorderColor = &H80000010
        lin(6).x1 = 0: lin(6).y1 = 0: lin(6).x2 = lin(6).x1: lin(6).y2 = Screen.Height: lin(6).BorderColor = &H8000000F
        lin(7).x1 = Screen.TwipsPerPixelX: lin(7).y1 = 0: lin(7).x2 = lin(7).x1: lin(7).y2 = Screen.Height: lin(7).BorderColor = &H8000000E
        
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
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width * 2
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
End Sub

Private Sub rtfSentence_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal node As MSComctlLib.node)
    If Val(Mid(node.Key, 2)) <> 0 Then
        Call ShowList(Val(Mid(node.Key, 2)))
    Else
        rtfSentence.Text = ""
        vsList.Rows = vsList.FixedRows
    End If
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset
    Dim rsValue As ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngStart As Long, strText As String
    
    If NewRow = OldRow Or NewRow < vsList.FixedRows Then Exit Sub
    
    On Error GoTo errH
    
    strSql = "Select 内容性质,内容文本,要素名称,要素单位 From 病历词句组成 Where 词句ID=[1] Order by 排列次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, Val(vsList.RowData(vsList.Row)))
    
    rtfSentence.Text = ""
    
    Do While Not rsTmp.EOF
        lngStart = Len(rtfSentence.Text)
        rtfSentence.SelStart = lngStart
        rtfSentence.SelLength = 0
        
        '自由文字
        strText = Nvl(rsTmp!内容文本)
        With rtfSentence
            .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
            .SelUnderline = False
        End With
       
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
