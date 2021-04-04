VERSION 5.00
Begin VB.Form frm供应商过滤 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件设置"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt信用额 
      Height          =   300
      Index           =   1
      Left            =   2565
      TabIndex        =   12
      Top             =   2235
      Width           =   1215
   End
   Begin VB.TextBox Txt信用额 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   10
      Top             =   2235
      Width           =   1215
   End
   Begin VB.TextBox Txt信用期 
      Height          =   300
      Index           =   1
      Left            =   2565
      TabIndex        =   8
      Top             =   1815
      Width           =   1215
   End
   Begin VB.TextBox TxtName 
      Height          =   300
      Left            =   915
      TabIndex        =   5
      Top             =   1380
      Width           =   2865
   End
   Begin VB.TextBox TxtCode 
      Height          =   300
      Index           =   1
      Left            =   2565
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox TxtCode 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame fra类型 
      Caption         =   "类型"
      Height          =   1770
      Left            =   3870
      TabIndex        =   22
      Top             =   795
      Width           =   1320
      Begin VB.CheckBox chkType 
         Caption         =   "卫材(&W)"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   24
         Tag             =   "4"
         Top             =   1485
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "其它(&Q)"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   16
         Tag             =   "4"
         Top             =   1170
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "药品(&Y)"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Tag             =   "1"
         Top             =   270
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "物资(&M)"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   14
         Tag             =   "2"
         Top             =   570
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox chkType 
         Caption         =   "设备(&S)"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   15
         Tag             =   "4"
         Top             =   870
         Value           =   1  'Checked
         Width           =   990
      End
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   20
      Top             =   705
      Width           =   5415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4110
      TabIndex        =   18
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2925
      TabIndex        =   17
      Top             =   2955
      Width           =   1100
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   19
      Top             =   2805
      Width           =   5415
   End
   Begin VB.TextBox Txt信用期 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   7
      Top             =   1815
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   2250
      TabIndex        =   11
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "信用额(&G)"
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Top             =   2295
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   2250
      TabIndex        =   23
      Top             =   1875
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "信用期(&X)"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   6
      Top             =   1875
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2250
      TabIndex        =   2
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "编码(&D)"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   1020
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "按下面的条件设置你所需要过滤内容。"
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Top             =   375
      Width           =   4110
   End
   Begin VB.Image img晋升 
      Height          =   480
      Left            =   195
      Picture         =   "frm供应商过滤.frx":0000
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frm供应商过滤"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnCancel As Boolean
Private mstrFilter As String
Dim mstrPrivs As String
Private Const mlngModule = 1025
Private mcllFilter As Collection

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intByte As Integer
    mblnCancel = False
    mstrFilter = ""
    'by lesfeng 2009-12-2 性能优化 信用期及信用额 过滤存在问题 现已经修改
    Set mcllFilter = New Collection
    mcllFilter.Add Array("", ""), "编码"
    mcllFilter.Add "", "名称"
    mcllFilter.Add Array("0", "0"), "信用期"
    mcllFilter.Add Array("0", "0"), "信用额"
    
    If Trim(TxtCode(0).Text) <> "" And Trim(TxtCode(1).Text) = "" Then
        mstrFilter = mstrFilter & " and 编码>=[1]"
    ElseIf Trim(TxtCode(1).Text) = "" And Trim(TxtCode(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and 编码<=[2]"
    ElseIf Trim(TxtCode(1).Text) <> "" And Trim(TxtCode(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and 编码>=[1] and 编码<=[2]"
    End If
    
    mcllFilter.Remove "编码"
    mcllFilter.Add Array(Trim(TxtCode(0).Text), Trim(TxtCode(1).Text)), "编码"
    
    If Trim(TxtName.Text) <> "" Then
        mstrFilter = mstrFilter & " and 名称 like [3]"
        mcllFilter.Remove "名称"
        mcllFilter.Add GetMatchingSting(TxtName.Text), "名称"
    End If
    
    If Trim(Txt信用期(0).Text) <> "" And Trim(Txt信用期(1).Text) = "" Then
        mstrFilter = mstrFilter & " and 信用期>=[4]"
    ElseIf Trim(Txt信用期(1).Text) = "" And Trim(Txt信用期(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and 信用期<=[5]"
    ElseIf Trim(Txt信用期(0).Text) <> "" And Trim(Txt信用期(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and 信用期>=[4] and 信用期<=[5]"
    End If
    mcllFilter.Remove "信用期"
    mcllFilter.Add Array(Val(Txt信用期(0).Text), Val(Txt信用期(1).Text)), "信用期"
    
    If Trim(Txt信用额(0).Text) <> "" And Trim(Txt信用额(1).Text) = "" Then
        mstrFilter = mstrFilter & " and 信用额>=[6]"
    ElseIf Trim(Txt信用额(1).Text) = "" And Trim(Txt信用额(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and 信用额<=[7]"
    ElseIf Trim(Txt信用额(0).Text) <> "" And Trim(Txt信用额(1).Text) <> "" Then
        mstrFilter = mstrFilter & " and 信用额>=[6] and 信用额<=[7]"
    End If
    mcllFilter.Remove "信用额"
    mcllFilter.Add Array(Val(Txt信用额(0).Text), Val(Txt信用额(1).Text)), "信用额"
    
    Dim i As Long
    Dim str类型 As String
    Dim strTmp As String
    
    str类型 = ""
    strTmp = ""
    For i = 0 To 4
        If chkType(i).Value = 1 And chkType(i).Enabled = True Then
            str类型 = str类型 & " or substr(类型," & i + 1 & ",1)=1"
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    
    Call zlDatabase.SetPara("供应商类型", strTmp, glngSys, mlngModule)
 
    If str类型 <> "" And str类型 <> "00000" Then
        str类型 = " And (" & Mid(str类型, 4) & ") "
    Else
        '所有
        str类型 = ""
    End If
    
    
    mstrFilter = mstrFilter & str类型
    If mstrFilter <> "" Then
        mstrFilter = Mid(mstrFilter, 5)
    End If
    Unload Me
End Sub


Private Sub TxtCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub TxtCode_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt信用期, KeyAscii, m数字式
End Sub

Private Sub TxtCode_LostFocus(Index As Integer)
    ImeLanguage False
End Sub

Private Sub TxtName_GotFocus()
    SetTxtGotFocus TxtName, True
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt信用期, KeyAscii, m文本式
End Sub

Private Sub TxtName_LostFocus()
    ImeLanguage False
End Sub

Private Sub Txt信用额_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub Txt信用额_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt信用期, KeyAscii, m金额式
End Sub

Private Sub Txt信用额_LostFocus(Index As Integer)
    ImeLanguage False
End Sub

Private Sub Txt信用期_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt信用期_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt信用期, KeyAscii, m金额式
End Sub

Public Sub GetFiler(ByVal FrmMain As Object, ByRef blnCancel As Boolean, ByRef strFilter As String, ByRef cllFilter As Collection, Optional ByVal strPriv As String)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取条件
    '--入参数:frmMain-主窗体
    '
    '--出参数:blnCancel-取消
    '         strFilter-条件
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim i As Integer
    mstrPrivs = strPriv
    strReg = zlDatabase.GetPara("供应商类型", glngSys, mlngModule)
    If strReg = "" Then
        strReg = "00000"
    End If
    Err = 0
    On Error Resume Next
    For i = 1 To Len(strReg)
        If Mid(strReg, i, 1) = 1 Then
            chkType(i - 1).Value = 1
        Else
            chkType(i - 1).Value = 0
        End If
    Next
    Call 权限控制
    Me.Show 1, FrmMain
    blnCancel = mblnCancel
    strFilter = mstrFilter
    Set cllFilter = mcllFilter
End Sub

Private Sub Txt信用期_LostFocus(Index As Integer)
    ImeLanguage False
End Sub


Private Sub 权限控制()
    '权限控制
    Dim bln药品 As Boolean
    Dim bln物资 As Boolean
    Dim bln设备 As Boolean
    Dim bln其他 As Boolean
    Dim bln卫材 As Boolean
    
    bln药品 = InStr(1, mstrPrivs, "药品供应商") <> 0
    bln物资 = InStr(1, mstrPrivs, "物资供应商") <> 0
    bln设备 = InStr(1, mstrPrivs, "设备供应商") <> 0
    bln其他 = InStr(1, mstrPrivs, "其他供应商") <> 0
    bln卫材 = InStr(1, mstrPrivs, "卫材供应商") <> 0
    
    chkType(0).Enabled = bln药品
    chkType(1).Enabled = bln物资
    chkType(2).Enabled = bln设备
    chkType(3).Enabled = bln其他
    chkType(4).Enabled = bln卫材
End Sub

