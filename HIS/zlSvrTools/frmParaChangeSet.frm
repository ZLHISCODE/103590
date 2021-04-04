VERSION 5.00
Begin VB.Form frmParaChangeSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "改变参数性质"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmParaChangeSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8745
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7500
      TabIndex        =   31
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Frame fra调整信息 
      Caption         =   "变动信息"
      Height          =   1050
      Left            =   0
      TabIndex        =   23
      Top             =   4260
      Width           =   8640
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   915
         TabIndex        =   25
         Top             =   255
         Width           =   7605
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   6
         Left            =   5025
         TabIndex        =   29
         Top             =   675
         Width           =   3495
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   5
         Left            =   915
         TabIndex        =   27
         Top             =   675
         Width           =   2415
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "变动原因"
         Height          =   180
         Index           =   11
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "变动时间"
         Height          =   180
         Index           =   10
         Left            =   4245
         TabIndex        =   28
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "变动人"
         Height          =   180
         Index           =   9
         Left            =   330
         TabIndex        =   26
         Top             =   675
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "授权方式变动"
      Height          =   1065
      Left            =   5040
      TabIndex        =   18
      Top             =   3105
      Width           =   3600
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   3
         Left            =   1275
         TabIndex        =   20
         Text            =   "需要授权"
         Top             =   315
         Width           =   1515
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   1
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   630
         Width           =   1995
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "原授权方式："
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
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "现授权方式："
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
         Index           =   7
         Left            =   105
         TabIndex        =   21
         Top             =   675
         Width           =   1170
      End
   End
   Begin VB.Frame fra参数变动 
      Caption         =   "参数类型变动"
      Height          =   2925
      Left            =   5040
      TabIndex        =   8
      Top             =   105
      Width           =   3585
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   4
         Left            =   1275
         TabIndex        =   10
         Text            =   "公共模块"
         Top             =   315
         Width           =   1350
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   0
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2070
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "参数类型说明："
         ForeColor       =   &H80000011&
         Height          =   180
         Index           =   4
         Left            =   75
         TabIndex        =   13
         Top             =   1020
         Width           =   1260
      End
      Begin VB.Label lblMemo 
         Caption         =   "  本机私有模块表示针对该模块分操作用户及分站点的参数。"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   3
         Left            =   60
         TabIndex        =   17
         Top             =   2460
         Width           =   3420
      End
      Begin VB.Label lblMemo 
         Caption         =   "  私有模块表示针对该模块分操作用户但不分站点的参数。"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   2040
         Width           =   3420
      End
      Begin VB.Label lblMemo 
         Caption         =   "  本机公共模块表示针对该模块不分操作用户但需要分机器的参数。"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   1635
         Width           =   3420
      End
      Begin VB.Label lblMemo 
         Caption         =   "  公共模块表示针对该模块不分操作用户不分机器的参数。"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   1245
         Width           =   3420
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "现参数类型："
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
         Left            =   120
         TabIndex        =   11
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "原参数类型："
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
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame fra参数 
      Caption         =   "参数基本信息"
      Height          =   4050
      Left            =   0
      TabIndex        =   30
      Top             =   105
      Width           =   4980
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   2
         Left            =   1095
         TabIndex        =   5
         Top             =   1125
         Width           =   3810
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   1
         Left            =   1095
         TabIndex        =   3
         Tag             =   "参数号"
         Top             =   765
         Width           =   960
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   0
         Left            =   1095
         TabIndex        =   1
         Tag             =   "模块"
         Top             =   375
         Width           =   3810
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         Caption         =   "   ddddddddddd"
         Height          =   2145
         Index           =   4
         Left            =   75
         TabIndex        =   7
         Tag             =   "参数说明"
         Top             =   1815
         Width           =   4815
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "参数说明："
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
         Index           =   3
         Left            =   75
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "参数名称："
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
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "参数号："
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
         Left            =   270
         TabIndex        =   2
         Top             =   765
         Width           =   780
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "模块："
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
         Index           =   0
         Left            =   450
         TabIndex        =   0
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6345
      TabIndex        =   32
      Top             =   5400
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaChangeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng参数ID As Long
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mstrUserName As String
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private Enum mTxt_idx
    idx_模块 = 0
    idx_参数号 = 1
    idx_参数名称 = 2
    idx_原授权方式 = 3
    idx_原参数类型 = 4
    idx_变动人 = 5
    idx_变动时间 = 6
    idx_变动原因 = 7
End Enum
Public Function ShowEdit(ByVal frmMain As Form, ByVal lng参数id As Long, ByVal strUserName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:显示编辑窗口
    '入参:frmMain-主窗体
    '     lng参数ID-参数值
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-19 14:16:42
    '-----------------------------------------------------------------------------------------------------------
    mlng参数ID = lng参数id: mblnChange = False: mblnOK = False: mstrUserName = strUserName: mblnFirst = True
    Me.Show 1
    ShowEdit = mblnOK
End Function
Private Function GetParaType(ByVal lng模块 As Long, ByVal int私有 As Integer, ByVal int本机 As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取参数类型
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-17 16:44:21
    '-----------------------------------------------------------------------------------------------------------
    If lng模块 = 0 Then
        '不存模块,证明只有两种类型:公共全局和私有全局
        GetParaType = IIf(int私有 = 0, "公共全局", "私有全局")
        Exit Function
    End If
    '对模块的处理
    If int本机 = 0 Then
        '不是本机的情况,只有两种类型:公共模块和私有模块
         GetParaType = IIf(int私有 = 0, "公共模块", "私有模块")
         Exit Function
    End If
    '对本机的模块进行处理也有两种情况:
    GetParaType = IIf(int私有 = 0, "本机公共模块", "本机私有模块")
End Function
Private Function LoadParaInfor() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载参数信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-19 14:19:26
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strCurDate As String
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.b_Public.Get_Current_Date")
    strCurDate = Format(rsTemp!日期, "yyyy-mm-dd HH:MM:SS")
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parameter", mlng参数ID)
    'ID, 系统,模块,私有,参数号,参数名, 参数值, 缺省值, 参数说明, 本机, 授权, 固定,模块名称
    If rsTemp.EOF Then
        MsgBox "参数未找到，可能已经被他人删除,请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
    mblnNotClick = True

    txtEdit(idx_模块) = Nvl(rsTemp!系统) & "-" & Nvl(rsTemp!模块)
    txtEdit(idx_参数号) = Nvl(rsTemp!参数号)
    txtEdit(idx_参数名称) = Nvl(rsTemp!参数名)
    lblEdit(4) = Nvl(rsTemp!影响控制说明)
    txtEdit(idx_变动人) = mstrUserName
    txtEdit(idx_变动时间) = strCurDate
    txtEdit(idx_原参数类型) = GetParaType(Val(Nvl(rsTemp!模块)), Val(Nvl(rsTemp!私有)), Val(Nvl(rsTemp!本机)))
    txtEdit(idx_原授权方式) = IIf(Val(Nvl(rsTemp!授权)) = 0, "不需要授权", "需要授权")
    txtEdit(idx_原授权方式).Tag = Val(Nvl(rsTemp!授权))
    If Val(Nvl(rsTemp!固定)) = 1 Then
        MsgBox "该参数为系统固定参数，不能调整！", vbOKOnly, gstrSysName
        mblnNotClick = False
        Exit Function
    End If
    mblnNotClick = False
    LoadParaInfor = True
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存相关的变动数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-19 15:01:37
    '-----------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo errHand:
    Dim int私有 As Integer, int本机 As Integer, int授权 As Integer
    Select Case cboEdit(0).Text
    Case "公共全局", "私有全局"
        MsgBox "注意:" & vbCrLf & _
               "   参数不能变动为公共全局和私有全局,请检查！", vbOKOnly, gstrSysName
        Exit Function
    Case "公共模块"
        int私有 = 0: int本机 = 0
    Case "私有模块"
        int私有 = 1: int本机 = 0
    Case "本机公共模块"
         int私有 = 0: int本机 = 1
    Case "本机私有模块"
         int私有 = 1: int本机 = 1
    Case ""
        MsgBox "注意:" & vbCrLf & _
               "   参数不能变动为空,请检查！", vbOKOnly, gstrSysName
        Exit Function
    End Select
    If cboEdit(1).Text = "不需要授权" Then
        int授权 = 0
    Else
        int授权 = 1
    End If
    SaveData = False
    'zl_Parameters_Change
    gstrSQL = "zl_Parameters_Change("
    '  参数id_In   zlParameters.ID%Type,
    gstrSQL = gstrSQL & "" & mlng参数ID & ","
    '  私有_In     zlParameters.私有%Type,
    gstrSQL = gstrSQL & "" & int私有 & ","
    '  本机_In     zlParameters.本机%Type,
    gstrSQL = gstrSQL & "" & int本机 & ","
    '  授权_In     zlParameters.授权%Type,
    gstrSQL = gstrSQL & "" & int授权 & ","
    '  变动人_In   Zlparachangedlog.变动人%Type,
    gstrSQL = gstrSQL & "'" & txtEdit(mTxt_idx.idx_变动人).Text & "',"
    '  变动原因_In Zlparachangedlog.变动原因%Type
    gstrSQL = gstrSQL & "'" & txtEdit(mTxt_idx.idx_变动原因).Text & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    SaveData = True
    Exit Function
errHand:
    MsgBox "注意:" & vbCrLf & _
           "   参数保存时发生错误，错误信息如下：" & vbCrLf & _
           "错误信息:" & err.Number & "-" & err.Description, vbOKOnly, gstrSysName
End Function
Private Sub InitCombox()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化Combox信息
    '编制:刘兴洪
    '日期:2009-02-19 14:43:55
    '-----------------------------------------------------------------------------------------------------------
    mblnNotClick = True
    With cboEdit(0)
        .AddItem "公共模块"
        .ItemData(.NewIndex) = 2
        If Trim(txtEdit(mTxt_idx.idx_原参数类型).Text) = "公共模块" Then .ListIndex = .NewIndex
        .AddItem "私有模块"
        .ItemData(.NewIndex) = 0
        If Trim(txtEdit(mTxt_idx.idx_原参数类型).Text) = "私有模块" Then .ListIndex = .NewIndex
        .AddItem "本机公共模块"
        .ItemData(.NewIndex) = 1
        If Trim(txtEdit(mTxt_idx.idx_原参数类型).Text) = "本机公共模块" Then .ListIndex = .NewIndex
        .AddItem "本机私有模块"
        .ItemData(.NewIndex) = 0
        If Trim(txtEdit(mTxt_idx.idx_原参数类型).Text) = "本机私有模块" Then .ListIndex = .NewIndex
    End With
    
    With cboEdit(1)
        .AddItem "不需要授权"
        If Trim(txtEdit(mTxt_idx.idx_原授权方式).Text) = "不需要授权" Then .ListIndex = .NewIndex
        .AddItem "需要授权"
        If Trim(txtEdit(mTxt_idx.idx_原授权方式).Text) = "需要授权" Then .ListIndex = .NewIndex
    End With
    mblnNotClick = False
End Sub
Private Function IsValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据的否法性
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-19 14:55:14
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parameter", mlng参数ID)
    'ID, 系统,模块,私有,参数号,参数名, 参数值, 缺省值, 参数说明, 本机, 授权, 固定,模块名称
    If rsTemp.EOF Then
        MsgBox "参数未找到，可能已经被他人删除,请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
    If Val(Nvl(rsTemp!固定)) = 1 Then
        MsgBox "该参数为系统固定参数，不能调整！", vbOKOnly, gstrSysName
        Exit Function
    End If
    If ActualLen(txtEdit(mTxt_idx.idx_变动原因).Text) > 200 Then
        MsgBox "变动原因最多能输入200个字符或100个汉字，不能调整！", vbOKOnly, gstrSysName
        txtEdit(mTxt_idx.idx_变动原因).SetFocus
        Exit Function
    End If
    If InStr(1, txtEdit(mTxt_idx.idx_变动原因).Text, "'") > 0 Then
        MsgBox "变动原因含有非法字符单引号，请检查！", vbOKOnly, gstrSysName
        txtEdit(mTxt_idx.idx_变动原因).SetFocus
        Exit Function
    End If
    IsValied = True
End Function

Private Sub SetCtlEnbaled()
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置相关控件属性
    '编制:刘兴洪
    '日期:2009-02-19 14:49:28
    '-----------------------------------------------------------------------------------------------------------
    Dim blnOk As Boolean
    mblnNotClick = True
    With cboEdit(0)
        Select Case .ItemData(.ListIndex)
        Case 1   '可以改变授权
            cboEdit(1).Enabled = True
        Case 2   '强制为授权
            cboEdit(1).Enabled = False
            cboEdit(1).ListIndex = 1
        Case Else   '不需授权
            cboEdit(1).Enabled = False
            cboEdit(1).ListIndex = 0
        End Select
    End With
    mblnNotClick = False
    
    blnOk = Trim(txtEdit(mTxt_idx.idx_原参数类型)) <> Trim(cboEdit(0).Text)
    blnOk = blnOk Or Trim(txtEdit(mTxt_idx.idx_原授权方式)) <> Trim(cboEdit(1).Text)
    cmdOK.Enabled = blnOk
End Sub
Private Sub cboEdit_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    Call SetCtlEnbaled
End Sub

Private Sub cmdClear_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If IsValied = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call InitCombox
    If LoadParaInfor() = False Then
        Unload Me: Exit Sub
    End If
    Call SetCtlEnbaled
    mblnChange = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    SendKeys "{tab}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mblnNotClick = True Then Exit Sub

    mblnChange = True
    Call SetCtlEnbaled
End Sub

Private Sub txtEdit_Click(Index As Integer)
    Call SetCtlEnbaled
    mblnChange = True
End Sub
