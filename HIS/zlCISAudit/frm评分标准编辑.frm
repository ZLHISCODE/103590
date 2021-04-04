VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm评分标准编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "评分标准编辑"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   Icon            =   "frm评分标准编辑.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSource 
      Caption         =   "数据源"
      Height          =   795
      Left            =   5190
      TabIndex        =   28
      Top             =   1005
      Visible         =   0   'False
      Width           =   2280
      Begin VB.OptionButton optSource 
         Caption         =   "标准版"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   30
         Top             =   210
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton optSource 
         Caption         =   "EMR库"
         Height          =   180
         Index           =   1
         Left            =   780
         TabIndex        =   29
         Top             =   525
         Width           =   1125
      End
   End
   Begin VB.ComboBox cmb否决等级 
      Height          =   300
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4185
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6540
      TabIndex        =   22
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5205
      TabIndex        =   21
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   330
      TabIndex        =   25
      Top             =   4860
      Width           =   1100
   End
   Begin VB.OptionButton opt录入方式 
      Caption         =   "按等级(&D)"
      Height          =   210
      Index           =   1
      Left            =   1530
      TabIndex        =   16
      Top             =   4230
      Width           =   1185
   End
   Begin VB.OptionButton opt录入方式 
      Caption         =   "按分数(&S)"
      Height          =   210
      Index           =   0
      Left            =   1530
      TabIndex        =   11
      Top             =   3795
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.TextBox txt标准分值 
      Height          =   300
      Left            =   3420
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3750
      Width           =   1185
   End
   Begin VB.ComboBox cmb缺陷等级 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frm评分标准编辑.frx":000C
      Left            =   3420
      List            =   "frm评分标准编辑.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4185
      Width           =   1185
   End
   Begin VB.ComboBox cmb评分单位 
      Height          =   300
      Left            =   6240
      TabIndex        =   15
      Top             =   3750
      Width           =   1185
   End
   Begin VB.TextBox txt描述 
      Height          =   705
      Left            =   1530
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1085
      Width           =   5940
   End
   Begin VB.TextBox txt方案名称 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      Top             =   215
      Width           =   5940
   End
   Begin VB.TextBox txt名称 
      Height          =   705
      Left            =   1530
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   1085
      Width           =   5940
   End
   Begin VB.CommandButton cmdXM 
      Caption         =   "…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   7155
      TabIndex        =   24
      Top             =   685
      Width           =   285
   End
   Begin VB.TextBox txt判断依据_NotCheck 
      Height          =   1320
      Left            =   1530
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   5940
   End
   Begin VB.CommandButton cmdCheck 
      Height          =   300
      Left            =   7470
      Picture         =   "frm评分标准编辑.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1910
      Width           =   300
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   1890
      Top             =   4500
      Width           =   3795
      _extentx        =   6694
      _extenty        =   741
      font            =   "frm评分标准编辑.frx":00EE
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2175
      Left            =   1515
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   5940
      _cx             =   10477
      _cy             =   3836
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
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
      BackColorBkg    =   16777215
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm评分标准编辑.frx":0116
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
   Begin VB.TextBox txt上级项目 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   650
      Width           =   5940
   End
   Begin VB.Label labNote 
      AutoSize        =   -1  'True
      Caption         =   "注："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1500
      TabIndex        =   27
      Top             =   3360
      Width           =   390
   End
   Begin VB.Label lab病人ID 
      AutoSize        =   -1  'True
      Caption         =   "[病人ID]、[主页ID]为系统参数，分别代表系统中的病人ID和主页ID。"
      Height          =   180
      Left            =   1950
      TabIndex        =   26
      Top             =   3360
      Width           =   5580
   End
   Begin VB.Label lab否决等级 
      Caption         =   "否决等级(&T)"
      Height          =   210
      Left            =   5205
      TabIndex        =   19
      Top             =   4260
      Width           =   1005
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8200
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   8200
      Y1              =   4680
      Y2              =   4665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "录入方式(&I)"
      Height          =   180
      Left            =   450
      TabIndex        =   10
      Top             =   3810
      Width           =   990
   End
   Begin VB.Label lblFS2 
      Caption         =   "评分单位(&W)"
      Height          =   210
      Left            =   5205
      TabIndex        =   14
      Top             =   3795
      Width           =   1005
   End
   Begin VB.Label lblDJ 
      Caption         =   "等级(&G)"
      Enabled         =   0   'False
      Height          =   210
      Left            =   2730
      TabIndex        =   17
      Top             =   4230
      Width           =   1005
   End
   Begin VB.Label lblFS1 
      Caption         =   "分数(&F)"
      Height          =   210
      Left            =   2730
      TabIndex        =   12
      Top             =   3795
      Width           =   1005
   End
   Begin VB.Label lbl描述 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目描述(&M)"
      Height          =   180
      Left            =   465
      TabIndex        =   6
      Top             =   1965
      Width           =   990
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "方案名称(&N)"
      Height          =   180
      Left            =   465
      TabIndex        =   0
      Top             =   270
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上级项目(&X)"
      Height          =   180
      Left            =   465
      TabIndex        =   2
      Top             =   675
      Width           =   990
   End
   Begin VB.Label lbl名称 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目名称(&B)"
      Height          =   180
      Left            =   465
      TabIndex        =   4
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "frm评分标准编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////
'       功能：能够新增评分标准（主项目、子项目），能够修改已有的评分标准。
'       吴庆伟  2005/1/6
'       注：如果该方案已经使用，则它下属的任何评分标准均不允许修改，也不能新增。
'///////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private m_lngID                 As Long     '当前编辑的评分标准的ID号
Private m_lng方案ID             As Long
Private m_lng上级ID             As Long
Private m_strEditMode           As String
Private m_lngOldRow             As Long
Private m_lngCurRow             As Long
Private m_blnModed              As Boolean
Private zlCheck                 As New clsCheck

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

'==============================================================================
'=功能：公共接口函数：用于传入初始化参数:ID '方式为插入，且ID存在，则在ID值前节点插入。
'==============================================================================
Public Sub ShowForm(方式 As String, 方案ID As Long, Optional 上级ID As Long = 0, Optional ID As Long = 0, Optional blnUsed As Boolean = True)
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo errH
    txt描述.Locked = Not blnUsed
    txt名称.Locked = Not blnUsed
    cmdXM.Enabled = blnUsed
    zlCheck.Sys_System Me
    
    m_blnModed = False
    m_lng方案ID = 方案ID  '必选参数
    m_lng上级ID = 上级ID
    m_lngID = ID          '为0表示新增
    m_lngCurRow = -1
    If m_lng方案ID < 1 Then
        Unload Me
        MsgBox "请先选择一个评分方案，如果你还没有录入方案请先录入！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call FillCmbs
    Call FillInitFixData
    
    m_strEditMode = 方式
    If m_strEditMode = "新增" Then
        Me.Caption = "新增" & IIf(上级ID = 0, "项目", "标准")
    ElseIf m_strEditMode = "插入" Then
        Me.Caption = "插入" & IIf(上级ID = 0, "项目", "标准")
    Else
        Me.Caption = "修改" & IIf(上级ID = 0, "项目", "标准")
        FillInitData
        txt上级项目.TabIndex = 0
    End If
    
    '如果已经有了下级项目，就不允许选择上级项目了！
    gstrSQL = "select count(*) from 病案评分标准视图 where 隐藏='是' and ID = [1] and 方案ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID, m_lng方案ID)
    If rsTemp.Fields(0) > 0 And m_strEditMode = "修改" Then
        '已有下级项目
        cmdXM.Enabled = False
    End If
    rsTemp.Close

    Me.Show 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：调整显示
'==============================================================================
Private Sub ShowSet()
    Dim bln一级项目         As Boolean

    On Error GoTo errH
    
    If gobjEmr Is Nothing Then
        fraSource.Visible = False
        optSource(0).Value = True
    Else
        fraSource.Visible = True
        optSource(0).Value = False
        optSource(1).Value = True
    End If
    
    If Len(txt上级项目) > 0 Then
        bln一级项目 = False
    Else
        bln一级项目 = True
    End If
    
    If bln一级项目 Then
        fraSource.Visible = False '项目无需指明数据源
        lbl描述.Visible = True
        lbl名称.Caption = "项目名称(&B)"
        lbl描述.Caption = "项目描述(&M)"
        txt名称.Visible = True
        txt描述.Move txt名称.Left, lbl描述.Top, txt名称.Width, 1500
        txt判断依据_NotCheck.Visible = False
        lab病人ID.Visible = False
        labNote.Visible = False
        cmdCheck.Visible = False
    Else
        cmdCheck.Visible = True
        lbl名称.Caption = "标准名称(&B)"
        lbl描述.Caption = "判断依据(&M)"
        txt名称.Visible = False
        txt描述.Visible = True
        txt判断依据_NotCheck.Visible = True
        lab病人ID.Visible = True
        labNote.Visible = True
        txt描述.Move txt名称.Left, txt名称.Top, IIf(gobjEmr Is Nothing, txt名称.Width, txt名称.Width - fraSource.Width - 100), txt名称.Height
        txt名称.Text = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：根据其他ID填入固定数据：如方案ID、上级ID
'==============================================================================
Private Sub FillInitFixData()
    Dim rs As ADODB.Recordset

    On Error GoTo errH
    
    gstrSQL = "select 名称 from 病案评分方案 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
    If Not rs.EOF Then
        txt方案名称 = IIf(IsNull(rs.Fields("名称")), "", rs.Fields("名称"))
    Else
        Unload Me
        MsgBox "请先选择评分方案。", vbInformation, "方案ID错误"
        Exit Sub
    End If
    gstrSQL = "select 名称 from 病案评分标准 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng上级ID)
    
    If Not rs.EOF Then
        txt上级项目 = IIf(IsNull(rs.Fields("名称")), "", rs.Fields("名称"))
    End If
    rs.Close
    
    Call ShowSet
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：根据ID填入初始数据
'==============================================================================
Private Sub FillInitData()
    Dim rs              As ADODB.Recordset
    On Error GoTo errH
    
    gstrSQL = "select A.ID,A.上级ID,A.方案ID,A.名称,A.描述,A.标准分值,A.缺陷等级,A.评分单位,A.上级序号,A.序号,A.判断依据,B.名称 as 上级项目,A.否决等级,A.数据源 from 病案评分标准 A,病案评分标准 B where A.上级ID=B.ID(+) and A.ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID)
    If Not rs.EOF Then
        txt上级项目.Text = IIf(IsNull(rs.Fields("上级项目")), "", rs.Fields("上级项目"))
        txt名称.Text = IIf(IsNull(rs.Fields("名称")), "", rs.Fields("名称"))
        txt描述.Text = IIf(IsNull(rs.Fields("描述")), "", rs.Fields("描述"))
        txt判断依据_NotCheck.Text = IIf(IsNull(rs.Fields("判断依据")), "", rs.Fields("判断依据"))
        If rs!数据源 = 0 Then
            optSource(0).Value = True
            optSource(1).Value = False
        Else
            optSource(0).Value = False
            optSource(1).Value = True
        End If
        If IsNull(rs.Fields("缺陷等级")) Then
            txt标准分值 = IIf(IsNull(rs.Fields("标准分值")), 0, IIf(rs.Fields("标准分值") < 1, Format(rs.Fields("标准分值"), "0.0"), rs.Fields("标准分值")))
            cmb评分单位.Text = IIf(IsNull(rs.Fields("评分单位")), "", rs.Fields("评分单位"))
            Set录入方式 0
        Else
            Select Case rs.Fields("缺陷等级")
                Case "甲"
                    cmb缺陷等级.ListIndex = 0
                Case "乙"
                    cmb缺陷等级.ListIndex = 1
                Case "丙"
                    cmb缺陷等级.ListIndex = 2
                Case "否"
                    cmb缺陷等级.ListIndex = 3
            End Select
            Select Case rs.Fields("否决等级")
                Case "乙"
                    cmb否决等级.ListIndex = 0
                Case "丙"
                    cmb否决等级.ListIndex = 1
                Case "不", ""
                    cmb否决等级.ListIndex = 2
            End Select
            Set录入方式 1
        End If
        zlControl.TxtSelAll txt上级项目
    Else
        Unload Me
        MsgBox "初始化数据错误，没有发现该条评分标准！请重试。", vbOKOnly + vbInformation, "参数错误"
        Exit Sub
    End If
    Call ShowSet
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：填入选择项的数据
'==============================================================================
Private Sub FillCmbs()
    On Error GoTo errH
    
    With cmb缺陷等级
        .AddItem "甲级"
        .AddItem "乙级"
        .AddItem "丙级"
        .AddItem "单项否决"
        .ListIndex = 1
    End With
    With cmb评分单位
        .AddItem ""
        .AddItem "项"
        .AddItem "处"
        .AddItem "个"
    End With
    With cmb否决等级
        .AddItem "乙级"
        .AddItem "丙级"
        .AddItem "不合格"
        .ListIndex = 1
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：评分单位提示
'==============================================================================
Private Sub cmb评分单位_GotFocus()
    On Error GoTo errH
    Call zlCommFun.OpenIme(True)
    ShowTips cmb评分单位, "请输入评分单位。", "评分单位"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmb缺陷等级_Click()
    cmb否决等级.Enabled = (cmb缺陷等级.Text = "单项否决")
End Sub

'==============================================================================
'=功能：缺陷等级提示
'==============================================================================
Private Sub cmb缺陷等级_GotFocus()
    On Error GoTo errH
    
    ShowTips cmb缺陷等级, "病案缺陷等级设定。", "缺陷等级"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：缺陷等级按回车确定就保存数据
'==============================================================================
Private Sub cmb缺陷等级_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmb缺陷等级_Validate(Cancel As Boolean)
    cmb否决等级.Enabled = (cmb缺陷等级.Text = "单项否决")
End Sub

'==============================================================================
'=功能：取消保存
'==============================================================================
Private Sub cmdCancel_Click()
    On Error GoTo errH
    m_blnModed = False
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 判断依据检测
'==============================================================================
Private Sub cmdCheck_Click()
    On Error GoTo errH
    Call CheckAuditSql_IN(Trim(txt判断依据_NotCheck.Text), True, IIf(optSource(0).Value, 0, 1))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：点击帮助
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo errH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：确定保存数据
'==============================================================================
Private Sub cmdOk_Click()
Dim blnBigXM        As Boolean
Dim strT As String, str等级 As String, strAudit As String, str否决等级 As String, intSource As Integer
    
    On Error GoTo errH
    
    '检查合法性
    If Not IsValid() Then Exit Sub
    intSource = IIf(optSource(0).Value, 0, 1)
    If m_lng上级ID <= 0 Then
       '上级项目为空，表示为一级项目
        blnBigXM = True
    Else
        blnBigXM = False
    End If
    Select Case cmb缺陷等级.Text
        Case "甲级"
            str等级 = "甲"
        Case "乙级"
            str等级 = "乙"
        Case "丙级"
            str等级 = "丙"
        Case "单项否决"
            str等级 = "否"
    End Select
    If str等级 = "否" Then
        Select Case cmb否决等级.Text
            Case "乙级"
                str否决等级 = "乙"
            Case "丙级"
                str否决等级 = "丙"
            Case "不合格"
                str否决等级 = "不"
        End Select
    End If
    '判断依据保存时需单引换双引
    strAudit = Replace(txt判断依据_NotCheck.Text, "'", "''")
    If m_strEditMode = "新增" Or m_strEditMode = "插入" Then
        If cmdCheck.Visible Then
            If Not CheckAuditSql_IN(Trim(txt判断依据_NotCheck.Text), False, IIf(optSource(0).Value, 0, 1)) Then Exit Sub
        End If
        If blnBigXM Then
            strT = "ZL_病案评分标准_Insert"
            gstrSQL = strT & _
                    "(" & zlDatabase.GetNextId("病案评分标准") & "," & IIf(m_lng上级ID = 0, "Null", CStr(m_lng上级ID)) & "," & m_lng方案ID & ",'" & txt名称 & "','" & txt描述 & _
                    "'," & IIf(txt标准分值.Enabled = False, "null", Val(txt标准分值)) & ",'" & IIf(opt录入方式(0).Value, "", str等级) & "','" & cmb评分单位.Text & "'," & m_lngID & ",'" & strAudit & "','" & str否决等级 & "'," & intSource & ")"
        Else
            strT = "ZL_病案评分标准_Insert"
            gstrSQL = strT & _
                    "(" & zlDatabase.GetNextId("病案评分标准") & "," & IIf(m_lng上级ID = 0, "Null", CStr(m_lng上级ID)) & "," & m_lng方案ID & ",'" & txt名称 & "','" & txt描述 & _
                    "'," & IIf(txt标准分值.Enabled = False, "null", Val(txt标准分值)) & ",'" & IIf(opt录入方式(0).Value, "", str等级) & "','" & cmb评分单位.Text & "',0,'" & strAudit & "','" & str否决等级 & "'," & intSource & ")"
        End If
    Else
        If cmdCheck.Visible Then
            If Not CheckAuditSql_IN(Trim(txt判断依据_NotCheck.Text), False, IIf(optSource(0).Value, 0, 1)) Then Exit Sub
        End If
        strT = "ZL_病案评分标准_Update"
        gstrSQL = strT & _
                "(" & CStr(m_lngID) & "," & IIf(m_lng上级ID = 0, "Null", CStr(m_lng上级ID)) & "," & m_lng方案ID & ",'" & txt名称 & "','" & txt描述 & _
                "'," & IIf(txt标准分值.Enabled = False, "null", Val(txt标准分值)) & ",'" & IIf(opt录入方式(0).Value, "", str等级) & "','" & cmb评分单位.Text & "','" & strAudit & "','" & str否决等级 & "'," & intSource & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    m_blnModed = True
    '手工刷新
    Call frm评分标准维护.DataLoad
    zlCheck.Msg_OK "评分标准保存成功！"
    If m_strEditMode = "新增" Then
        txt名称.Text = ""
        txt描述.Text = ""
        txt标准分值.Text = ""
        cmb评分单位.Text = ""
        txt判断依据_NotCheck.Text = ""
        opt录入方式(0).Value = True
        If txt名称.Visible = True Then
            txt名称.SetFocus
        Else
            txt描述.SetFocus
        End If
    Else
        Unload Me
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：数据检测正确与否
'=返回：有效返回True,否则为False
'=说明：同一方案的项目名称不能重复
'==============================================================================
Private Function IsValid() As Boolean
    Dim bln一级项目         As Boolean
    
    On Error GoTo errH
    
    IsValid = False
    If Len(txt上级项目) > 0 Then
        bln一级项目 = False
    Else
        bln一级项目 = True
    End If
    If bln一级项目 And m_strEditMode = "新增" Then
        Dim rsTmp As New ADODB.Recordset
        gstrSQL = "select 名称 from 病案评分标准 where 上级ID is null and 方案ID = [1] And 名称 = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID, Trim(txt名称.Text))
        If Not rsTmp.EOF Then
            zlCheck.Msg_OK "同一方案中的项目名称不能重复！"
            zlControl.TxtSelAll txt名称: txt名称.SetFocus
            Exit Function
        End If
    End If
    '调用StrIsValid函数来确保字符串格式正确，注意：长度使用的是lenB值（对应数据表定义中的值）
    If zlCommFun.StrIsValid(txt名称.Text, txt名称.MaxLength * 2) = False Then
        zlCheck.Msg_OK "请输入名称！"
        zlControl.TxtSelAll txt名称: txt名称.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txt描述.Text, txt描述.MaxLength * 2) = False Then
        zlCheck.Msg_OK "请输入描述！"
        zlControl.TxtSelAll txt描述: txt描述.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(cmb评分单位.Text, 8) = False Then
        zlCheck.Msg_OK "评分单位长度超过4个汉字！请重新录入。"
        cmb评分单位.SetFocus
        Exit Function
    End If
    If Len(Trim(txt名称)) = 0 And Len(Trim(txt描述)) = 0 Then
        zlCheck.Msg_OK "名称和描述不能同时为空！请重新录入。"
        If txt名称.Visible = True Then
            zlControl.TxtSelAll txt名称: txt名称.SetFocus
        Else
            zlControl.TxtSelAll txt描述: txt描述.SetFocus
        End If
        Exit Function
    End If
    If opt录入方式(0).Value And Len(Trim(txt标准分值)) = 0 Then
        zlCheck.Msg_OK "请输入标准分值！"
        zlControl.TxtSelAll txt标准分值: txt标准分值.SetFocus
        Exit Function
    End If
    If Len(Trim(txt标准分值)) > 0 Then
        If Not IsNumeric(txt标准分值) Then
            zlCheck.Msg_OK "请输入标准分值！"
            zlControl.TxtSelAll txt标准分值: txt标准分值.SetFocus
            Exit Function
        End If
        If Val(txt标准分值.Text) > 9999# Then
            zlCheck.Msg_OK "输入的标准分值太大！"
            zlControl.TxtSelAll txt标准分值: txt标准分值.SetFocus
            Exit Function
        End If
    End If
    
    '当前工作站没有安装新版病历组件，但却在修改数据源为EMR库的判断依据，应禁止
    If gobjEmr Is Nothing And optSource(1).Value Then
        zlCheck.Msg_OK "当前工作站未安装病历组件，禁止修改需要在EMR库执行的判断依据！"
        Exit Function
    End If
    
    IsValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能：动态装入一级项目
'==============================================================================
Private Sub cmdXM_Click()
    Dim rsTemp              As ADODB.Recordset
    
    On Error GoTo errH

    If cmdXM.Tag = "打开" Then
        cmdXM.Tag = ""
        Grid.Visible = False
        Grid_SelChange
        Exit Sub
    Else
        cmdXM.Tag = "打开"
    End If
    
    With Grid
        .Clear
        .Redraw = flexRDNone
        If m_strEditMode = "修改" Then
            '如果是编辑模式，一级项目需要排除它本身
            gstrSQL = "select ID,标准分值,名称,描述 from 病案评分标准 Where 上级ID is null and ID <> [1] and 方案ID = [2]"
        Else
            gstrSQL = "select ID,标准分值,名称,描述 from 病案评分标准 Where 上级ID is null and 方案ID = [2]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID, m_lng方案ID)
        
        .FocusRect = flexFocusSolid
        '数据填入
        .Cols = 4
        .Rows = rsTemp.RecordCount + 2
        Dim i As Long
        .Cell(flexcpText, 0, 0) = "ID"
        .Cell(flexcpText, 0, 1) = "名称"
        .Cell(flexcpText, 0, 2) = "标准分值"
        .Cell(flexcpText, 0, 3) = "描述"
        .Cell(flexcpText, 1, 0) = "0"
        .Cell(flexcpText, 1, 1) = "<空>"
        .Cell(flexcpText, 1, 2) = "<空>"
        .Cell(flexcpText, 1, 3) = "<空>"
        i = 2
        Do Until rsTemp.EOF
            If m_lng上级ID = rsTemp.Fields("ID") Then m_lngCurRow = i
            .Cell(flexcpText, i, 0) = IIf(IsNull(rsTemp.Fields("ID")), "", rsTemp.Fields("ID"))
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("名称")), "", rsTemp.Fields("名称"))
            .Cell(flexcpText, i, 2) = IIf(IsNull(rsTemp.Fields("标准分值")), "", Format(rsTemp.Fields("标准分值"), "####分"))
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("描述")), "", IIf(Len(rsTemp.Fields("描述")) > 30, Left(rsTemp.Fields("描述"), 27) + "...", rsTemp.Fields("描述")))
            rsTemp.MoveNext
            i = i + 1
        Loop
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 250
        .RowHeightMax = 300
        .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("标准分值")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("名称")) = flexAlignLeftTop
        .ColAlignment(.ColIndex("描述")) = flexAlignLeftTop
        '宽度设置
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("名称")) = 1600
        .ColWidth(.ColIndex("标准分值")) = 800
        .ColWidth(.ColIndex("描述")) = 4000
        '最大宽度设置
        .ColWidthMax = 4000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1
        .AutoSize 3
        .SelectionMode = flexSelectionByRow
        .AllowBigSelection = False
        
        '选中先前的行
        If m_lngCurRow > 0 And m_lngCurRow < i Then
            .Row = m_lngCurRow
            .ShowCell m_lngCurRow, 2
        Else
            m_lngCurRow = 1
            .Row = 1
            .ShowCell m_lngCurRow, 2
        End If
        .ZOrder 0
        .Redraw = flexRDBuffered
        .Visible = True
        If .Visible = True Then .SetFocus
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：项目得到焦点时显示Tips提示
'==============================================================================
Private Sub cmdXM_GotFocus()
    On Error GoTo errH
    ShowTips cmdXM, "点击或回车将弹出选择下拉框，用于选择上级项目。", "上级项目"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：项目按F1显示Tips提示
'==============================================================================
Private Sub cmdXM_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = vbKeyF1 Then
        ShowTips cmdXM, "选择上级项目请点击该按钮。", "选择上级项目"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面控件初始化
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    Call InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面初始化
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：项目选择状态
'==============================================================================
Private Sub Grid_Click()
    On Error GoTo errH
    Call Grid_SelChange
    If cmdXM.Tag = "打开" Then
        cmdXM.Tag = ""
        Grid.Visible = False
        Exit Sub
    Else
        cmdXM.Tag = "打开"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：Grid得到焦点时显示Tips提示
'==============================================================================
Private Sub Grid_GotFocus()
    On Error GoTo errH
    ShowTips cmdXM, "点击或回车将选定该上级项目。", "选定项目"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：Grid按键确认
'==============================================================================
Private Sub Grid_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt名称.Visible = True Then
            zlControl.TxtSelAll txt名称
            txt名称.SetFocus
        Else
            zlControl.TxtSelAll txt描述
            txt描述.SetFocus
        End If
        Grid.Visible = False
        cmdXM.Tag = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：Grid焦点行号及ID
'==============================================================================
Private Sub Grid_SelChange()
    Dim m_lngID As Long

    On Error GoTo errH
    
    m_lng上级ID = 0
    m_lngCurRow = Grid.Row
    If m_lngCurRow <= 0 Then Exit Sub
    m_lngCurRow = Grid.Row     '获取行号
    m_lng上级ID = Grid.Cell(flexcpText, m_lngCurRow, 0)
    txt上级项目.Text = IIf(m_lng上级ID = 0, "", Grid.Cell(flexcpText, m_lngCurRow, 1))
    m_lngOldRow = m_lngCurRow
    Call ShowSet
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub optSource_Click(Index As Integer)
    If Index = 0 Then
        lab病人ID.Caption = "[病人ID]、[主页ID]为系统参数，分别代表系统中的病人ID和主页ID。"
        txt判断依据_NotCheck.Tag = 0 '用于放大编辑窗口
    Else
        lab病人ID.Caption = "[MID]、[ALIDIN]为系统参数，分别代表EMR中的病人ID和入院ID"
        lab病人ID.ToolTipText = "使用EMR中的ID,除[MID]、[ALIDIN]外都需要使用HextoRaw转换，以便用到索引。"
        txt判断依据_NotCheck.Tag = 1 '用于放大编辑窗口
    End If
End Sub
'==============================================================================
'=功能：录入方式设定
'==============================================================================
Private Sub opt录入方式_Click(Index As Integer)
    On Error GoTo errH
    Set录入方式 Index
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：根据录入方式设置控件状态
'==============================================================================
Private Sub Set录入方式(i As Integer)
    On Error GoTo errH
    If i = 0 Then
        cmb缺陷等级.Enabled = False
        lblDJ.Enabled = False
        txt标准分值.Enabled = True
        lblFS1.Enabled = True
        lblFS2.Enabled = True
        cmb评分单位.Enabled = True
        cmb否决等级.Enabled = False
    Else
        cmb缺陷等级.Enabled = True
        lblDJ.Enabled = True
        txt标准分值.Enabled = False
        lblFS1.Enabled = False
        lblFS2.Enabled = False
        cmb评分单位.Enabled = False
        cmb否决等级.Enabled = (cmb缺陷等级.Text = "单项否决")
    End If
    opt录入方式(i).Value = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：录入方式得到焦点显示Tips提示
'==============================================================================
Private Sub opt录入方式_GotFocus(Index As Integer)
    On Error GoTo errH
    If Index = 0 Then
        ShowTips opt录入方式(0), "按照打分的方式进行评分，此处提供该项的标准分数（一律输入正数）。加分与扣分由分制来确定。", "按分数"
    Else
        ShowTips opt录入方式(1), "对于某些重要评分项，如果不合格那么整个病案直接定位“乙级”或“丙级”，需要在这里选择其缺陷等级。“单项否决”表示如果该项不合格则整个病案不合格，不纳入等级评定。", "按等级"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：txt标准分值 得到焦点显示Tips提示
'==============================================================================
Private Sub txt标准分值_GotFocus()
    On Error GoTo errH
    ShowTips txt标准分值, "请输入正数。", "标准分值"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：txt方案名称 得到焦点显示Tips提示
'==============================================================================
Private Sub txt方案名称_GotFocus()
    On Error GoTo errH
    ShowTips txt方案名称, "评分方案的名称。在这里不允许修改。", "方案名称"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：txt描述 得到焦点显示Tips提示
'==============================================================================
Private Sub txt描述_GotFocus()
    On Error GoTo errH
    ShowTips txt描述, "用于录入主评分项目的描述或者子评分项目的名称。", "描述"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：txt名称 得到焦点显示Tips提示
'==============================================================================
Private Sub txt名称_GotFocus()
    On Error GoTo errH
    zlControl.TxtSelAll txt名称
    ShowTips txt名称, "如果不是子评分项目，可以在这里输入主评分项目的名称。", "名称"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：txt上级项目 得到焦点显示Tips提示
'==============================================================================
Private Sub txt上级项目_GotFocus()
    On Error GoTo errH
    zlControl.TxtSelAll txt上级项目
    ShowTips txt上级项目, "如果你需要创建子评分项目，请先选择它的上级项目。", "上级项目"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：控件通用Tips显示
'==============================================================================
Private Sub ShowTips(ctl As Control, str内容 As String, Optional str标题 As String = "提示信息", Optional lng时间 As Long = 2500)
    Dim X As Single, Y As Single
    On Error GoTo errH
    X = (ctl.Left + ctl.Width / 2) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height) / Screen.TwipsPerPixelY
    If Len(str内容) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        tipPopup1.TimeOut = lng时间
        tipPopup1.Title = str标题
        tipPopup1.Text = str内容
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



