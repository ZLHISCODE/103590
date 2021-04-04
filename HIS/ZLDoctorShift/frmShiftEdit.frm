VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmShiftEdit 
   BackColor       =   &H80000004&
   Caption         =   "病人交接班内容编辑"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14940
   Icon            =   "frmShiftEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14940
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSplitX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6615
      ScaleWidth      =   45
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox picMainBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   8415
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   600
      Width           =   8415
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   240
         ScaleHeight     =   3225
         ScaleWidth      =   7965
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   7995
         Begin VB.PictureBox picEdit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3975
            Left            =   120
            ScaleHeight     =   3975
            ScaleWidth      =   7815
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   120
            Width           =   7815
            Begin VB.PictureBox picPanel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1260
               Left            =   600
               ScaleHeight     =   1260
               ScaleWidth      =   7125
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   120
               Width           =   7125
               Begin VB.CommandButton cmdType 
                  BackColor       =   &H80000004&
                  Caption         =   "…"
                  Height          =   250
                  Left            =   6600
                  TabIndex        =   21
                  TabStop         =   0   'False
                  ToolTipText     =   "选择(*)"
                  Top             =   0
                  Width           =   270
               End
               Begin VB.CommandButton cmdFind 
                  Caption         =   "…"
                  Height          =   270
                  Left            =   3000
                  TabIndex        =   22
                  TabStop         =   0   'False
                  ToolTipText     =   "查找当前科室的正在住院的病人"
                  Top             =   0
                  Width           =   270
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   7
                  Left            =   5025
                  Locked          =   -1  'True
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   1960
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   6
                  Left            =   3375
                  Locked          =   -1  'True
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   735
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   5
                  Left            =   855
                  Locked          =   -1  'True
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   1575
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   4
                  Left            =   5040
                  Locked          =   -1  'True
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   440
                  Width           =   1920
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   3
                  Left            =   3345
                  Locked          =   -1  'True
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   440
                  Width           =   735
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   2
                  Left            =   855
                  Locked          =   -1  'True
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Top             =   457
                  Width           =   1560
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  Height          =   290
                  Index           =   1
                  Left            =   855
                  TabIndex        =   1
                  Top             =   10
                  Width           =   2400
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   0
                  Left            =   4320
                  Locked          =   -1  'True
                  TabIndex        =   2
                  TabStop         =   0   'False
                  Top             =   10
                  Width           =   2520
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "入院时间"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   7
                  Left            =   4260
                  TabIndex        =   30
                  Top             =   975
                  Width           =   720
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "入院途径"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   6
                  Left            =   2565
                  TabIndex        =   29
                  Top             =   975
                  Width           =   720
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "住院号"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   5
                  Left            =   240
                  TabIndex        =   28
                  Top             =   975
                  Width           =   540
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "床号"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   4
                  Left            =   4635
                  TabIndex        =   27
                  Top             =   500
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "年龄"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   26
                  Top             =   500
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "性别"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   2
                  Left            =   390
                  TabIndex        =   25
                  Top             =   500
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "姓名"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   1
                  Left            =   420
                  TabIndex        =   24
                  Top             =   60
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "病人类型"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   23
                  Top             =   60
                  Width           =   720
               End
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   4920
               TabIndex        =   19
               Top             =   2400
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "多选"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   4560
               TabIndex        =   18
               Top             =   2400
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               Height          =   290
               Index           =   0
               Left            =   2880
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   2400
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   290
               Index           =   0
               Left            =   5880
               ScaleHeight     =   285
               ScaleWidth      =   1335
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   2160
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "目前诊断"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   31
               Top             =   2400
               Visible         =   0   'False
               Width           =   720
            End
         End
         Begin VB.VScrollBar vscBar 
            Height          =   7575
            LargeChange     =   200
            Left            =   7800
            SmallChange     =   200
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   8640
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.Frame fraInfo 
         BackColor       =   &H80000004&
         Caption         =   "交班描述"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2415
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4575
         Begin RichTextLib.RichTextBox rtbBox 
            Height          =   1935
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   3413
            _Version        =   393217
            BackColor       =   -2147483644
            BorderStyle     =   0
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmShiftEdit.frx":6852
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin zlSubclass.Subclass Subclass 
      Left            =   1200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label lblWdith 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "宽度计算"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   9120
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   720
   End
   Begin XtremeCommandBars.CommandBars cbsExec 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmShiftEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WM_MOUSEWHEEL = &H20A          '鼠标滚动
Private Const con表格线X = 470  'SBAR表格线X
Private Const con间距 = 75  '间距

Private Enum MenuType
        ID_新增 = 1
        ID_修改 = 2
        ID_删除 = 3
        ID_保存 = 4
        ID_取消 = 5
        ID_病案 = 6
        
        ID_类型 = 99
End Enum

Private Enum PatiInfo
        idx病人类型 = 0
        idx姓名 = 1
        idx性别 = 2
        idx年龄 = 3
        idx床号 = 4
        idx住院号 = 5
        idx入院途径 = 6
        idx入院时间 = 7
End Enum

Private Type Ctl布局
    'X
    对齐线1 As Long
    对齐线2 As Long
    
    'Y
    S区域线 As Long
    B区域线 As Long
    A区域线 As Long
    R区域线 As Long
    
    'Y
    lngTop As Long '动态控件顶点高度
    LngY As Long '用于布局时记录当前设置高度
End Type

Private Type InfoD
    '病人信息
    病人类型 As String '新入,一护...
    病人ID As Long
    主页ID As Long

    '交班记录
    交班记录ID As Long
    交班科室ID As Long
    交班开始时间 As Date
    交班结束时间 As Date
    
    '内容记录
    EditType As Long '编辑方式 0-预览　1-新增  2-修改
    内容ID As Long
    内容序号 As Long
End Type


Private mCtl布局 As Ctl布局 '布局信息
Private mInfo As InfoD '窗体变量

Private mobjCtl As Object '用于设置焦点

Private mbtnNoEdit As Boolean '不允许编辑
Private mblnChange As Boolean '控件记录改变
Private mblnSave As Boolean '内容是否已保存
Private mblnEnter As Boolean '是否回车
Private mblnLoad As Boolean '是否正在初始化
Private mbln预览 As Boolean

Public gstr预览类型 As String '预览窗体调用

'参数缓存
Private mstrLike As String  '是支持全匹配
Private mbln血库系统 As Boolean   '是否安装血库系统


'对象缓存
Public gfrmParent As Object             '父窗体对象
Private mrsCtlType As ADODB.Recordset   '病人类型记录集
Private mrsCtlInfo As ADODB.Recordset   '控件信息记录集

'字符串缓存
Private mstrCtls As String          '缓存控件信息字符串
Private mstrTextCtl As String       '文本框idx字符串
Private mstrPicCtl As String        '父控件idx字符串
Private mstrPatiType As String      '病人类型字符串


Private Sub chkInfo_Click(Index As Integer)
    If (mInfo.EditType <> 0 And mblnLoad = False) Or mbln预览 Then
        mblnChange = True
        Call SetTextVisible(Val(Split(chkInfo(Index).Tag, ",")(0)))
        Call MakeText
    End If
End Sub


Private Sub SetTextVisible(lngIndex As Long)
   '设置选择控件的文本框是否显示
    Dim i As Long, intType As Integer, blnVisible As Boolean

    On Error GoTo errH
    If lngIndex = 0 Then Exit Sub
    If InStr("," & mstrTextCtl & ",", "," & lngIndex & ",") = 0 Or InStr("," & mstrPicCtl & ",", "," & lngIndex & ",") = 0 Then Exit Sub
    intType = Val(Split(picTmp(lngIndex).Tag, ",")(0))
    
    If intType = 2 Then
            For i = 1 To optInfo.Count - 1
                If Val(Split(optInfo(i).Tag, ",")(0)) = lngIndex Then
                    If optInfo(i).Value = True And Split(optInfo(i).Tag, ",")(1) = "1" Then
                        blnVisible = True
                        Exit For
                    End If
                End If
            Next
    ElseIf intType = 3 Then
        For i = 1 To chkInfo.Count - 1
            If Val(Split(chkInfo(i).Tag, ",")(0)) = lngIndex Then
                If chkInfo(i).Value = 1 And Split(chkInfo(i).Tag, ",")(1) = "1" Then
                    blnVisible = True
                    Exit For
                End If
            End If
        Next
    End If
        
    If txtInfo(lngIndex).Visible <> blnVisible Then
         txtInfo(lngIndex).Visible = blnVisible
         txtInfo(lngIndex).Text = ""
         Call picMain_Resize
         Call DrawLine
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub chkInfo_GotFocus(Index As Integer)
    If mblnEnter Then Call ShowCtl(chkInfo(Index)): mblnEnter = False
End Sub

Private Sub cmdFind_Click()
    Call GetPatiList(1)
End Sub

Private Sub cmdType_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    On Error GoTo errH

    vPoint = zlcontrol.GetCoordPos(txtPatiInfo(idx病人类型).Container.hWnd, txtPatiInfo(idx病人类型).Left, txtPatiInfo(idx病人类型).Top)
    blnCancel = True
    
    strSQL = "Select a.顺序 As ID, a.简称, a.名称 ,Decode(b.C2, Null, 0, 1) As 已勾选check" & vbNewLine & _
            "From 医生交接班病人类型 A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.是否停用 = 0 And a.简称 = b.C2(+)" & vbNewLine & _
            "Order By 顺序"


    Set mobjCtl = txtPatiInfo(idx病人类型)
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择病人类型", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtPatiInfo(idx病人类型).Height, blnCancel, True, True, txtPatiInfo(idx病人类型).Text)
    
    zlcontrol.ControlSetFocus txtPatiInfo(idx病人类型)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!简称
                rsTmp.MoveNext
            Loop
            
            If mInfo.病人类型 <> Mid(strTmp, 2) Then
                Set mobjCtl = txtPatiInfo(idx病人类型)
                If mblnChange And txtPatiInfo(idx病人类型).Text <> "" Then
                    If MsgBox("切换病人类型将清除当前未保存的项目,请确认是否继续？", vbInformation + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
                        zlcontrol.ControlSetFocus txtPatiInfo(idx病人类型)
                        Exit Sub
                    End If
                End If
                
                txtPatiInfo(idx病人类型).Text = Mid(strTmp, 2)
                mInfo.病人类型 = txtPatiInfo(idx病人类型).Text
                Call UnloadCtl
                Call InitCtl
                Call IntData
                Call MakeText
            End If
        Else
            Set mobjCtl = txtPatiInfo(idx病人类型)
            MsgBox "未查找到可以选择的病人类型!", vbInformation, Me.Caption
            zlcontrol.ControlSetFocus txtPatiInfo(idx病人类型)
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Not mobjCtl Is Nothing Then
        zlcontrol.ControlSetFocus mobjCtl
        Set mobjCtl = Nothing
    End If
    
    
    '首次进入预览界面时界面排版概率出现错乱，进行容错处理
    If gstr预览类型 <> "" Then
        Call picMain_Resize
        Call Form_Resize
        gstr预览类型 = ""
    End If
End Sub

Private Sub Form_Load()

    On Error GoTo errH
    
    If gstr预览类型 <> "" Then mbln预览 = True
    
    '获取控件信息记录集
    Call GetCtlRs
    
    '处理按钮
    Call InitExecBar

    '固定布局
    picEdit.Top = 0: picEdit.Left = 0: picEdit.Width = picMain.Width: picEdit.Height = 10000
    picPanel.Top = 120: picPanel.Left = 600
    mCtl布局.lngTop = picPanel.Top + picPanel.Height
    
    '初始化布局
    picSplitX.BackColor = Me.BackColor
    picSplitX.Left = 8720
    
    '滚轮事件初始化
    Subclass.hWnd = Me.hWnd
    Subclass.Messages(WM_MOUSEWHEEL) = True
    
    '初始化参数
    mstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    mbln血库系统 = (IsSysSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    
    '界面初始化
    Call zlRefresh(0, 0, 0, 0, 0, Now - 1, Now + 1, False, gstr预览类型)

    '设置预览界面
    If mbln预览 Then
        txtPatiInfo(idx病人类型).Text = gstr预览类型
        Call SetPicEnabled(True)
        cmdType.Visible = False
        Call RestoreWinState(Me, App.ProductName)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Paint()
    Call DrawLine
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    
    On Error Resume Next
    
    lngTop = IIf(mbln预览, 500, 340)

    picSplitX.Top = lngTop: picSplitX.Height = Me.Height - lngTop - IIf(mbln预览, 250, 0)
    picMainBack.Move 0, lngTop, picSplitX.Left, Me.Height - lngTop - IIf(mbln预览, 600, 0)
    picMain.Move 0, 85, picMainBack.Width, picMainBack.Height - 85
    picInfo.Move picSplitX.Left + picSplitX.Width, lngTop, Me.Width - (picSplitX.Left + picSplitX.Width) - IIf(mbln预览, 270, 0), Me.Height - lngTop - IIf(mbln预览, 570, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjCtl = Nothing
    Set mrsCtlType = Nothing
    Set mrsCtlInfo = Nothing
    gstr预览类型 = ""
    
    Subclass.Messages(WM_MOUSEWHEEL) = False

    If mbln预览 Then Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl


    On Error GoTo errH
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        If mbln预览 = True Then
            .SetIconSize False, 24, 24
        Else
            .SetIconSize False, 16, 16
        End If
    End With
    Set cbsExec.Icons = zlCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        If mbln预览 Then
            '工具栏
            Set objControl = .Add(xtpControlButton, ID_类型, "病人类型：")
            objControl.IconId = 807

            If Not mrsCtlType Is Nothing Then
                mrsCtlType.Filter = ""
                Do While Not mrsCtlType.EOF
                    Set objControl = .Add(xtpControlButton, ID_类型 + Val(mrsCtlType!顺序 & ""), mrsCtlType!简称 & "")
                    objControl.IconId = IIf(InStr("," & gstr预览类型 & ",", "," & mrsCtlType!简称 & ",") > 0, 12, 10)

                    mrsCtlType.MoveNext
                Loop
            End If
        Else
            Set objControl = .Add(xtpControlButton, ID_病案, "电子病案查阅")
                objControl.IconId = 816
            Set objControl = .Add(xtpControlButton, ID_新增, "新增")
                objControl.IconId = 4112
                objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_修改, "修改")
                objControl.IconId = 4113
            Set objControl = .Add(xtpControlButton, ID_删除, "删除")
                objControl.IconId = 4114
            Set objControl = .Add(xtpControlButton, ID_保存, "保存")
                objControl.IconId = 3091
                objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_取消, "取消")
                objControl.IconId = 3014
        End If

    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub UnloadCtl()
    '卸载界面的控件
    Dim obj As Object
    
    On Error Resume Next
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In txtInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In optInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In chkInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In picTmp
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
End Sub

Private Sub CtlVisible(blnVisible As Boolean)
    '控制界面的控件显示
    Dim obj As Object
    
    On Error Resume Next
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In txtInfo
        If obj.Index <> 0 And InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") = 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In optInfo
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In chkInfo
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In picTmp
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
End Sub

Private Sub InitCtl()
    '初始化界面的动态控件
    Dim i As Long, j As Long, m As Long, n As Long
    Dim intIndex As Integer, intTabIndex As Integer, intOptIndex As Integer, intChkIndex As Integer
    Dim rsInfoCopy As ADODB.Recordset
    Dim arrTmp As Variant
    Dim blnCheck As Boolean, blnText As Boolean
    
    On Error GoTo errH
    
    '清空缓存项
    mstrCtls = "": mstrTextCtl = "": mstrPicCtl = ""
    
    If mrsCtlType Is Nothing Or mrsCtlInfo Is Nothing Or mInfo.病人类型 = "" Then
        Call picMain_Resize
        Exit Sub
    End If
    
    '还原记录集
    mrsCtlType.Filter = ""
    mrsCtlInfo.Filter = ""
    
    If mrsCtlType.EOF Or mrsCtlInfo.EOF Then Exit Sub
    
    mrsCtlType.MoveFirst
    mrsCtlInfo.MoveFirst
    
    Set rsInfoCopy = zlDatabase.CopyNewRec(mrsCtlInfo) '复制一个记录集，用于判断是否两列
    
    mblnLoad = True
    intTabIndex = 9
    For m = 1 To 4 '按SBAR的顺序加载
        mrsCtlType.MoveFirst
        For i = 1 To mrsCtlType.RecordCount
            If mInfo.病人类型 = "" Or InStr("," & mInfo.病人类型 & ",", "," & mrsCtlType!简称 & ",") > 0 Then
                    mrsCtlInfo.Filter = "病人简称='" & mrsCtlType!简称 & "' And 项目类别='" & Decode(m, 1, "S", 2, "B", 3, "A", 4, "R") & "'"
                    For j = 1 To mrsCtlInfo.RecordCount
                        
                        '检查加载顺序
                        rsInfoCopy.Filter = "项目名称 ='" & mrsCtlInfo!项目名称 & "'"
                        blnCheck = True
                        Do While Not rsInfoCopy.EOF
                            If rsInfoCopy!病人简称 & "" <> mrsCtlInfo!病人简称 & "" And InStr("," & mInfo.病人类型 & ",", "," & rsInfoCopy!病人简称 & ",") > 0 Then
                                If InStr(mstrPatiType, "," & rsInfoCopy!病人简称 & ",") < InStr(mstrPatiType, "," & mrsCtlInfo!病人简称 & ",") Then
                                    blnCheck = False
                                    Exit Do
                                End If
                            End If
                            rsInfoCopy.MoveNext
                        Loop
                        rsInfoCopy.Filter = ""
                        
                        '检查死亡则隐藏
                        If blnCheck Then
                            If ((InStr("," & mInfo.病人类型 & ",", "死亡") > 0 Or mInfo.病人类型 = "") And Val(mrsCtlInfo!死亡则隐藏 & "") = 1) Then
                                blnCheck = False
                            End If
                        End If
                    
                        If InStr(mstrCtls, "," & mrsCtlInfo!项目名称) = 0 And blnCheck Then
                        
                            intIndex = lblInfo.Count
                            Load lblInfo(intIndex)
                            
                            lblInfo(intIndex).Caption = mrsCtlInfo!项目名称 & ""
                            
                             '处理长度过长导致换行
                            If lblInfo(intIndex).Width > 1200 Then
                                lblInfo(intIndex).Caption = Mid(mrsCtlInfo!项目名称 & "", 1, Len(lblInfo(intIndex).Caption) / 2) & vbCrLf & Mid(mrsCtlInfo!项目名称 & "", Len(lblInfo(intIndex).Caption) / 2 + 1)
                            End If
                            
                            
                            '缓存控件一行排列个数
                            rsInfoCopy.Filter = "病人简称='" & mrsCtlType!简称 & "' And 项目类别='" & Decode(m, 1, "S", 2, "B", 3, "A", 4, "R") & "' And 序号=" & mrsCtlInfo!序号 & " And 项目名称<> '" & mrsCtlInfo!项目名称 & "'"
                            If Not rsInfoCopy.EOF Then
                                If InStr(mstrCtls, mrsCtlType!简称 & "," & rsInfoCopy!项目名称) = 0 Then
                                    lblInfo(intIndex).Tag = "1," & mrsCtlInfo!病人简称 & "," & mrsCtlInfo!项目类别 & "," & mrsCtlInfo!输入类型
                                Else
                                    lblInfo(intIndex).Tag = "2," & mrsCtlInfo!病人简称 & "," & mrsCtlInfo!项目类别 & "," & mrsCtlInfo!输入类型
                                End If
                            Else
                                If Val(mrsCtlInfo!输入形式 & "") = 1 And (Val(mrsCtlInfo!输入类型 & "") = 1 Or Val(mrsCtlInfo!输入类型 & "") = 2) Then
                                    lblInfo(intIndex).Tag = "1," & mrsCtlInfo!病人简称 & "," & mrsCtlInfo!项目类别 & "," & mrsCtlInfo!输入类型
                                Else
                                    lblInfo(intIndex).Tag = "0," & mrsCtlInfo!病人简称 & "," & mrsCtlInfo!项目类别 & "," & mrsCtlInfo!输入类型
                                End If
                            End If
                            
                            
                            Select Case Val(mrsCtlInfo!输入形式 & "")
                                    Case 1 '输入框
                                        Load txtInfo(intIndex)
                                        
                                        '控件Tab顺序处理
                                        intTabIndex = intTabIndex + 1
                                        txtInfo(intIndex).TabIndex = intTabIndex
                                        txtInfo(intIndex).Locked = Val(mrsCtlInfo!是否只读 & "") = 1
                                        txtInfo(intIndex).BackColor = IIf(txtInfo(intIndex).Locked, &H8000000F, &H80000005)
                                        txtInfo(intIndex).TabStop = Not txtInfo(intIndex).Locked
                                        
                                        '控件高度
                                        txtInfo(intIndex).Height = txtInfo(intIndex).Height * IIf(Val(mrsCtlInfo!输入行数 & "") = 0, 1, Val(mrsCtlInfo!输入行数 & ""))
                                        
                                        mstrTextCtl = mstrTextCtl & "," & intIndex
                                    Case 2 '单项选择
                                        Load picTmp(intIndex) '生成容器控件
                                        blnText = False
                                        mstrPicCtl = mstrPicCtl & "," & intIndex
                                        
                                        picTmp(intIndex).Tag = "2," & optInfo.Count '缓存项目对应选项控件开始
                                        arrTmp = Split(mrsCtlInfo!输入值域 & "", ",")
                                        For n = 0 To UBound(arrTmp)
                                            intOptIndex = optInfo.Count
                                            Load optInfo(intOptIndex)
                                            
                                            intTabIndex = intTabIndex + 1
                                            optInfo(intOptIndex).TabStop = True
                                            optInfo(intOptIndex).TabIndex = intTabIndex
                                            optInfo(intOptIndex).Caption = IIf(Mid(arrTmp(n), 1, 1) = "*", Mid(arrTmp(n), 2), arrTmp(n))
                                            optInfo(intOptIndex).Tag = intIndex & "," & IIf(Mid(arrTmp(n), 1, 1) = "*", "1", "")
                                            
                                            '计算选项宽度
                                            lblWdith.Caption = arrTmp(n)
                                            optInfo(intOptIndex).Width = lblWdith.Width + 300
                                      
                                            '设置父项容器
                                            Set optInfo(intOptIndex).Container = picTmp(intIndex)
                                            
                                            If Mid(arrTmp(n), 1, 1) = "*" Then blnText = True '存在文本框
                                        Next
                                        optInfo(Val(Split(picTmp(intIndex).Tag, ",")(1))).Value = True '暂时默认第一个为默认选项
                                        
                                        '设置带附加信息的文本框
                                        If blnText Then
                                            Load txtInfo(intIndex)
                                            '控件Tab顺序处理
                                            intTabIndex = intTabIndex + 1
                                            txtInfo(intIndex).TabIndex = intTabIndex
                                            txtInfo(intIndex).TabStop = True
                                            
                                            txtInfo(intIndex).Visible = Split(optInfo(Val(Split(picTmp(intIndex).Tag, ",")(1))).Tag, ",")(1) = "1"
                                            
                                            '控件高度
                                            txtInfo(intIndex).Height = txtInfo(intIndex).Height * IIf(Val(mrsCtlInfo!输入行数 & "") = 0, 1, Val(mrsCtlInfo!输入行数 & ""))
                                            
                                            mstrTextCtl = mstrTextCtl & "," & intIndex
                                        End If

                                        picTmp(intIndex).Tag = picTmp(intIndex).Tag & "," & optInfo.Count - 1 '缓存项目对应选项控件结束
                                    Case 3 '多项选择
                                        blnText = False
                                        Load picTmp(intIndex) '生成容器控件
                                        
                                        mstrPicCtl = mstrPicCtl & "," & intIndex
                                        
                                        picTmp(intIndex).Tag = "3," & chkInfo.Count '缓存项目对应选项控件开始
                                            
                                        arrTmp = Split(mrsCtlInfo!输入值域 & "", ",")
                                        For n = 0 To UBound(arrTmp)
                                            intChkIndex = chkInfo.Count
                                            Load chkInfo(intChkIndex)
                                            
                                            intTabIndex = intTabIndex + 1
                                            chkInfo(intChkIndex).TabStop = True
                                            chkInfo(intChkIndex).TabIndex = intTabIndex
                                            chkInfo(intChkIndex).Caption = IIf(Mid(arrTmp(n), 1, 1) = "*", Mid(arrTmp(n), 2), arrTmp(n))
                                            chkInfo(intChkIndex).Tag = intIndex & "," & IIf(Mid(arrTmp(n), 1, 1) = "*", "1", "")
                                            
                                            '计算选项宽度
                                            lblWdith.Caption = arrTmp(n)
                                            chkInfo(intChkIndex).Width = lblWdith.Width + 300
                                            
                                            '设置父项容器
                                            Set chkInfo(intChkIndex).Container = picTmp(intIndex)
                                            
                                            If Mid(arrTmp(n), 1, 1) = "*" Then blnText = True '存在文本框
                                            
                                            picTmp(intIndex).Refresh
                                        Next
                                        
                                        '设置带附加信息的文本框
                                        If blnText Then
                                            Load txtInfo(intIndex)
                                            '控件Tab顺序处理
                                            intTabIndex = intTabIndex + 1
                                            txtInfo(intIndex).TabIndex = intTabIndex
                                            txtInfo(intIndex).TabStop = True
                                            txtInfo(intIndex).Visible = False '默认为不显示
                                            
                                            '控件高度
                                            txtInfo(intIndex).Height = txtInfo(intIndex).Height * IIf(Val(mrsCtlInfo!输入行数 & "") = 0, 1, Val(mrsCtlInfo!输入行数 & ""))
                                            
                                            mstrTextCtl = mstrTextCtl & "," & intIndex
                                        End If

                                        picTmp(intIndex).Tag = picTmp(intIndex).Tag & "," & chkInfo.Count - 1 '缓存项目对应选项控件结束
                            End Select
                            
                            mstrCtls = mstrCtls & ";" & mrsCtlType!简称 & "," & mrsCtlInfo!项目名称
                            
                        End If
                        mrsCtlInfo.MoveNext
                    Next
            End If
            mrsCtlType.MoveNext
        Next
    Next
    
     
    Call picMain_Resize
    
    Call CtlVisible(True) '最后显示控件
    
    mblnLoad = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub CtlResize()
    '动态控件排版
    Dim i As Long, j As Long
    Dim lngTmp1 As Long, lngTmp2 As Long
    Dim lngTop As Long, lngMaxTop As Long
    Dim lng误差 As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngLeft As Long

    On Error GoTo errH

    On Error Resume Next
    
    '获取最大的标签宽度用于对齐
    For i = 1 To lblInfo.Count - 1
        If Val(Mid(lblInfo(i).Tag, 1, 1)) = 2 Then
            If lblInfo(i).Width > lngTmp2 Then lngTmp2 = lblInfo(i).Width
        Else
            If lblInfo(i).Width > lngTmp1 Then lngTmp1 = lblInfo(i).Width
        End If
    Next
    
    '计算对齐线
    If lngTmp1 = 0 Then lngTmp1 = lblPatiInfo(idx住院号).Width
    If lngTmp2 = 0 Then lngTmp2 = lblPatiInfo(idx病人类型).Width
    mCtl布局.对齐线1 = picPanel.Left + lngTmp1
    mCtl布局.对齐线2 = picPanel.Left + picPanel.Width / 2 + lngTmp2 + 200
    
     
    '处理病人信息控件对齐
    lngLeft = picPanel.Width / 3
'    '第一排
    lblPatiInfo(idx姓名).Left = mCtl布局.对齐线1 - picPanel.Left - lblPatiInfo(idx姓名).Width
    txtPatiInfo(idx姓名).Left = mCtl布局.对齐线1 - picPanel.Left + con间距: txtPatiInfo(idx姓名).Width = picPanel.Width / 2 - txtPatiInfo(idx姓名).Left - 10
    lblPatiInfo(idx病人类型).Left = mCtl布局.对齐线2 - picPanel.Left - lblPatiInfo(idx病人类型).Width
    txtPatiInfo(idx病人类型).Left = mCtl布局.对齐线2 - picPanel.Left + con间距: txtPatiInfo(idx病人类型).Width = picPanel.Width - txtPatiInfo(idx病人类型).Left - 10
    
    cmdFind.Top = txtPatiInfo(idx姓名).Top + 20: cmdFind.Left = txtPatiInfo(idx姓名).Left + txtPatiInfo(idx姓名).Width - cmdFind.Width - 15
    
    cmdType.Top = txtPatiInfo(idx病人类型).Top + 20: cmdType.Left = txtPatiInfo(idx病人类型).Left + txtPatiInfo(idx病人类型).Width - cmdType.Width - 15
    
    '第二排
    lblPatiInfo(idx性别).Left = mCtl布局.对齐线1 - picPanel.Left - lblPatiInfo(idx性别).Width
    txtPatiInfo(idx性别).Left = lblPatiInfo(idx性别).Left + lblPatiInfo(idx性别).Width + con间距: txtPatiInfo(idx性别).Width = lngLeft - txtPatiInfo(idx性别).Left - 10
    lblPatiInfo(idx年龄).Left = lngLeft + lblPatiInfo(idx入院途径).Width + 100 - lblPatiInfo(idx年龄).Width
    txtPatiInfo(idx年龄).Left = lblPatiInfo(idx年龄).Left + lblPatiInfo(idx年龄).Width + con间距: txtPatiInfo(idx年龄).Width = lngLeft * 2 - txtPatiInfo(idx年龄).Left - 10
    lblPatiInfo(idx床号).Left = lngLeft * 2 + lblPatiInfo(idx入院时间).Width + 100 - lblPatiInfo(idx床号).Width
    txtPatiInfo(idx床号).Left = lblPatiInfo(idx床号).Left + lblPatiInfo(idx床号).Width + con间距: txtPatiInfo(idx床号).Width = picPanel.Width - txtPatiInfo(idx床号).Left - 10

    '第三排
    lblPatiInfo(idx住院号).Left = mCtl布局.对齐线1 - picPanel.Left - lblPatiInfo(idx住院号).Width
    txtPatiInfo(idx住院号).Left = lblPatiInfo(idx住院号).Left + lblPatiInfo(idx住院号).Width + con间距: txtPatiInfo(idx住院号).Width = lngLeft - txtPatiInfo(idx住院号).Left - 10
    lblPatiInfo(idx入院途径).Left = lngLeft + 100
    txtPatiInfo(idx入院途径).Left = lblPatiInfo(idx入院途径).Left + lblPatiInfo(idx入院途径).Width + con间距: txtPatiInfo(idx入院途径).Width = lngLeft * 2 - txtPatiInfo(idx入院途径).Left - 10
    lblPatiInfo(idx入院时间).Left = lngLeft * 2 + 100
    txtPatiInfo(idx入院时间).Left = lblPatiInfo(idx入院时间).Left + lblPatiInfo(idx入院时间).Width + con间距: txtPatiInfo(idx入院时间).Width = picPanel.Width - txtPatiInfo(idx入院时间).Left - 10

    '默认布局变量
    mCtl布局.S区域线 = mCtl布局.lngTop + 80
    mCtl布局.B区域线 = 0
    mCtl布局.A区域线 = 0
    mCtl布局.R区域线 = 0
    mCtl布局.LngY = 0
    
    '设置标签位置
    For i = 1 To lblInfo.Count - 1
        lng误差 = 55
        If lblInfo(i).Height > lblPatiInfo(idx姓名).Height Then
            If InStr("," & mstrPicCtl & ",", "," & i & ",") = 0 And InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 Then
                If txtInfo(i).Height < lblInfo(i).Height Then
                    lng误差 = txtInfo(i).Height / 2 - lblInfo(i).Height / 2
                End If
            End If
        End If
        
        If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
            lng误差 = 25
        End If
        

        If Val(Mid(lblInfo(i).Tag, 1, 1)) = 2 Then '设置第二排
            lblInfo(i).Top = lblInfo(i - 1).Top
            lblInfo(i).Left = mCtl布局.对齐线2 - lblInfo(i).Width
            
            '存在选择框
            If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
                picTmp(i).Top = lblInfo(i).Top - lng误差: picTmp(i).Left = mCtl布局.对齐线2 + con间距
                picTmp(i).Width = picPanel.Left + picPanel.Width - picTmp(i).Left: picTmp(i).Height = 290
                
                
                lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
                lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
                
                '自适应计算选项框
                If Val(Mid(picTmp(i).Tag, 1, 1)) = 2 Then
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            optInfo(j).Top = 0: optInfo(j).Left = 0
                        Else
                            '当大于区域宽度时，自动换行
                            If optInfo(j - 1).Left + optInfo(j - 1).Width + 50 + optInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + optInfo(j).Height
                                optInfo(j).Top = optInfo(j - 1).Top + optInfo(j - 1).Height + 50: optInfo(j).Left = 0
                            Else
                                optInfo(j).Top = optInfo(j - 1).Top: optInfo(j).Left = optInfo(j - 1).Left + optInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                Else
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            chkInfo(j).Top = 0: chkInfo(j).Left = 0
                        Else
                            '当大于区域宽度时，自动换行
                            If chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50 + chkInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + chkInfo(j).Height
                                chkInfo(j).Top = chkInfo(j - 1).Top + chkInfo(j - 1).Height + 50: chkInfo(j).Left = 0
                            Else
                                chkInfo(j).Top = chkInfo(j - 1).Top: chkInfo(j).Left = chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                End If
                
                
                 If lngMaxTop < picTmp(i).Top + picTmp(i).Height Then lngMaxTop = picTmp(i).Top + picTmp(i).Height
                 If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                
                '存在文本框
                If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And txtInfo(i).Visible Then
                    txtInfo(i).Top = picTmp(i).Top + picTmp(i).Height + 20: txtInfo(i).Left = picTmp(i).Left
                    txtInfo(i).Width = picTmp(i).Width
                    
                    If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                    If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                End If
                
                
            End If
            
            '存在文本框
            If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And InStr("," & mstrPicCtl & ",", "," & i & ",") = 0 Then
                txtInfo(i).Top = lblInfo(i).Top - lng误差: txtInfo(i).Left = mCtl布局.对齐线2 + con间距
                txtInfo(i).Width = picPanel.Left + picPanel.Width - txtInfo(i).Left
                
                If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
            End If
            
            '计算区域线
            Select Case Split(lblInfo(i).Tag, ",")(2)
                Case "S"
                    mCtl布局.S区域线 = lngMaxTop + 150
                Case "B"
                    mCtl布局.B区域线 = lngMaxTop + 150
                Case "A"
                    mCtl布局.A区域线 = lngMaxTop + 150
                Case "R"
                    mCtl布局.R区域线 = lngMaxTop + 150
            End Select
        Else '设置第一排
        
            '动态布局高度计算
            If i = 1 Then
                lngTop = mCtl布局.lngTop + IIf(Split(lblInfo(i).Tag, ",")(2) = "S", 200, 300)
            Else
                lngTop = lngMaxTop + 200
            End If
            
            If i <> 1 Then
                If Split(lblInfo(i - 1).Tag, ",")(2) <> Split(lblInfo(i).Tag, ",")(2) Then
                    lngTop = lngMaxTop + 300
                End If
            End If

            lblInfo(i).Top = lngTop
            lblInfo(i).Left = mCtl布局.对齐线1 - lblInfo(i).Width
            
            '存在选择框
            If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
                picTmp(i).Top = lblInfo(i).Top - lng误差: picTmp(i).Left = mCtl布局.对齐线1 + con间距
                picTmp(i).Width = (picPanel.Left + picPanel.Width) / Decode(Val(Mid(lblInfo(i).Tag, 1, 1)), 0, 1, 2) - picTmp(i).Left
                picTmp(i).Height = 290
                
                lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
                lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
                
                '自适应计算选项框
                If Val(Mid(picTmp(i).Tag, 1, 1)) = 2 Then
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            optInfo(j).Top = 0: optInfo(j).Left = 0
                        Else
                            '当大于区域宽度时，自动换行
                            If optInfo(j - 1).Left + optInfo(j - 1).Width + 50 + optInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + optInfo(j).Height
                                optInfo(j).Top = optInfo(j - 1).Top + optInfo(j - 1).Height + 50: optInfo(j).Left = 0
                            Else
                                optInfo(j).Top = optInfo(j - 1).Top: optInfo(j).Left = optInfo(j - 1).Left + optInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                Else
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            chkInfo(j).Top = 0: chkInfo(j).Left = 0
                        Else
                            '当大于区域宽度时，自动换行
                            If chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50 + chkInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + chkInfo(j).Height
                                chkInfo(j).Top = chkInfo(j - 1).Top + chkInfo(j - 1).Height + 50: chkInfo(j).Left = 0
                            Else
                                chkInfo(j).Top = chkInfo(j - 1).Top: chkInfo(j).Left = chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                End If
                
                
                 If lngMaxTop < picTmp(i).Top + picTmp(i).Height Then lngMaxTop = picTmp(i).Top + picTmp(i).Height
                 If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                
                '存在文本框
                If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And txtInfo(i).Visible Then
                    txtInfo(i).Top = picTmp(i).Top + picTmp(i).Height + 20: txtInfo(i).Left = picTmp(i).Left
                    txtInfo(i).Width = picTmp(i).Width
                    
                    If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                    If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                End If
                
                
            End If
            
            '存在文本框
            If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And InStr("," & mstrPicCtl & ",", "," & i & ",") = 0 Then
                txtInfo(i).Top = lblInfo(i).Top - lng误差: txtInfo(i).Left = mCtl布局.对齐线1 + con间距
                txtInfo(i).Width = (picPanel.Left + picPanel.Width) / Decode(Val(Mid(lblInfo(i).Tag, 1, 1)), 0, 1, 2) - txtInfo(i).Left
                
                If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
            End If
            
            '计算区域线
            Select Case Split(lblInfo(i).Tag, ",")(2)
                Case "S"
                    mCtl布局.S区域线 = lngMaxTop + 150
                Case "B"
                    mCtl布局.B区域线 = lngMaxTop + 150
                Case "A"
                    mCtl布局.A区域线 = lngMaxTop + 150
                Case "R"
                    mCtl布局.R区域线 = lngMaxTop + 150
            End Select
        End If
    Next
    
    mCtl布局.LngY = lngMaxTop + 200

    Call DrawLine
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub optInfo_Click(Index As Integer)
    If (mInfo.EditType <> 0 And mblnLoad = False) Or mbln预览 Then
        mblnChange = True
        Call SetTextVisible(Val(Split(optInfo(Index).Tag, ",")(0)))
        Call MakeText
    End If
End Sub

Private Sub optInfo_GotFocus(Index As Integer)
    If mblnEnter Then Call ShowCtl(optInfo(Index)): mblnEnter = False
End Sub

Private Sub picEdit_Resize()
    On Error Resume Next
    picPanel.Width = picEdit.Width - picPanel.Left - 150
End Sub

Private Sub picMain_Resize()
    Call RefreshResize
End Sub

Public Sub RefreshResize()
    On Error Resume Next
    vscBar.Top = 0: vscBar.Height = picMain.Height
    vscBar.Left = picMain.Width - vscBar.Width
    

    '计算容器高度
    If picEdit.Height < picMain.Height Then picEdit.Height = picMain.Height
    vscBar.Visible = picMain.Height < mCtl布局.LngY
    picEdit.Width = IIf(vscBar.Visible, picMain.Width - vscBar.Width, picMain.Width)
    Call CtlResize
    picEdit.Height = mCtl布局.LngY
    If picEdit.Height < picMain.Height Then picEdit.Height = picMain.Height
    
    '切换滚动条显示时重新刷新
    If IIf(vscBar.Visible, 1, 0) <> IIf(picMain.Height < mCtl布局.LngY, 1, 0) Then
        vscBar.Visible = picMain.Height < picEdit.Height
        picEdit.Width = IIf(vscBar.Visible, picMain.Width - vscBar.Width, picMain.Width)
        Call CtlResize
        picEdit.Height = mCtl布局.LngY
        If picEdit.Height < picMain.Height Then picEdit.Height = picMain.Height
    End If
    
    vscBar.Max = picEdit.Height - picMain.Height
End Sub


Private Sub picInfo_GotFocus()
    On Error Resume Next
    If zlcontrol.IsCtrlSetFocus(txtPatiInfo(idx姓名)) Then
        txtPatiInfo(idx姓名).SetFocus
    Else
        Call SeekNextCtl
    End If
End Sub

Private Sub rtbBox_Change()
    If rtbBox.Visible And rtbBox.Locked = False And mInfo.EditType <> 0 And mblnLoad = False Then mblnChange = True
End Sub

Private Sub rtbBox_KeyPress(KeyAscii As Integer)
    If InStr("&'<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub


Private Sub rtbBox_LostFocus()
    If rtbBox.Locked = False And rtbBox.Text <> rtbBox.Tag Then
        With rtbBox
            .SelStart = 0
            .SelLength = Len(rtbBox.Text)
            .SelColor = RGB(30, 144, 255)
            .SelStart = Len(rtbBox.Text)
        End With
    End If
End Sub

Private Sub txtPatiInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> idx姓名 Then
        Call zlCommFun.ShowTipInfo(txtPatiInfo(Index).hWnd, txtPatiInfo(Index).Text, True, True)
    End If
End Sub

Private Sub txtPatiInfo_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo errH
    If Index = idx姓名 Then
        If txtPatiInfo(Index).Visible And txtPatiInfo(Index).Locked = False Then
            txtPatiInfo(Index).Text = txtPatiInfo(Index).Tag
        End If
    End If
    
    If txtPatiInfo(Index).Enabled = True And txtPatiInfo(Index).Locked = False Then
        Call MakeText
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vscBar_Change()
    On Error Resume Next
    picEdit.Top = -vscBar.Value
    picEdit.SetFocus
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Top = 0: fraInfo.Left = 0: fraInfo.Height = picInfo.Height - 25: fraInfo.Width = picInfo.Width - 25
    rtbBox.Height = fraInfo.Height - 300
    rtbBox.Width = fraInfo.Width - 200
End Sub


Private Sub picEdit_Paint()
    Call DrawLine
End Sub

Private Sub DrawLine()
    Dim lngUp As Long
    '画框框
    On Error Resume Next

    picEdit.Cls
    picEdit.Line (con表格线X, 0)-(con表格线X, picEdit.Height)
    picEdit.ForeColor = RGB(105, 105, 105)
    
    '病人类型为空时画SBAR线
    If mInfo.病人类型 = "" And mCtl布局.S区域线 > 0 Then
        mCtl布局.B区域线 = mCtl布局.S区域线 + (picEdit.Height - mCtl布局.S区域线) / 3
        mCtl布局.A区域线 = mCtl布局.S区域线 + ((picEdit.Height - mCtl布局.S区域线) / 3) * 2
        mCtl布局.R区域线 = mCtl布局.S区域线 + ((picEdit.Height - mCtl布局.S区域线) / 3) * 2
    End If
    
    If mCtl布局.S区域线 > 0 Then
        If mCtl布局.R区域线 = 0 And mCtl布局.A区域线 = 0 And mCtl布局.B区域线 = 0 Then
            picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
            picEdit.Print "S"
        Else
            lngUp = mCtl布局.S区域线
            picEdit.Line (0, mCtl布局.S区域线)-(picEdit.Width, mCtl布局.S区域线)
            picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = (mCtl布局.S区域线) / 2 - 100
            picEdit.Print "S"
        End If
    End If
    
    If mCtl布局.B区域线 > 0 Then
        If mCtl布局.R区域线 = 0 And mCtl布局.A区域线 = 0 Then
            picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
            picEdit.Print "B"
        Else
            picEdit.Line (0, mCtl布局.B区域线)-(picEdit.Width, mCtl布局.B区域线)
            
            picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = mCtl布局.B区域线 - (mCtl布局.B区域线 - lngUp) / 2 - 100
            picEdit.Print "B"
            
            lngUp = mCtl布局.B区域线
        End If
    End If
    
    If mCtl布局.A区域线 > 0 Then
        If mCtl布局.R区域线 = 0 Then
            picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
            picEdit.Print "A"
            
        Else
            picEdit.Line (0, mCtl布局.A区域线)-(picEdit.Width, mCtl布局.A区域线)
            
            picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = mCtl布局.A区域线 - (mCtl布局.A区域线 - lngUp) / 2 - 100
            picEdit.Print "A"
            lngUp = mCtl布局.A区域线
        End If
    End If
    
    If mCtl布局.R区域线 > 0 Then
        picEdit.FontName = "宋体": picEdit.FontSize = 13: picEdit.FontBold = True
        picEdit.CurrentX = 150
        picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
        picEdit.Print "R"
    End If
    
    picEdit.Refresh
End Sub


Private Sub GetCtlRs()
    '功能：获取控件信息记录集
    Dim strSQL As String

    On Error GoTo errH
    '病人类型记录集
    strSQL = "Select a.简称, a.名称, a.顺序, a.起始描述, a.提取sql From 医生交接班病人类型 A Where a.是否停用 = 0 order by A.顺序"
    Set mrsCtlType = zlDatabase.OpenSQLRecord(strSQL, "GetCtlRs")
    
    '缓存病人类型字符串
    If Not mrsCtlType Is Nothing Then
        mstrPatiType = ""
        Do While Not mrsCtlType.EOF
            mstrPatiType = mstrPatiType & "," & mrsCtlType!简称
            mrsCtlType.MoveNext
        Loop
        mstrPatiType = mstrPatiType & ","
        mrsCtlType.MoveFirst
    End If
        
    '控件信息记录集
    strSQL = "select 病人简称, 项目名称, 序号, 项目类别, 输入形式, 输入类型, 输入格式, 输入值域, 输入行数, 提取来源, 提取病历, 提取SQL, 描述文字, 是否只读, 死亡则隐藏 from 医生交接班病人项目  order by 病人简称,项目类别,序号,Rownum"
    Set mrsCtlInfo = zlDatabase.OpenSQLRecord(strSQL, "GetCtlRs")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If txtInfo(Index).Visible And txtInfo(Index).Locked = False And mInfo.EditType <> 0 And mblnLoad = False Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    zlcontrol.TxtSelAll txtInfo(Index)
    If mblnEnter Then Call ShowCtl(txtInfo(Index)): mblnEnter = False
End Sub


Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strMask As String
    On Error GoTo errH
    
    If InStr("&'<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextCtl
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        Select Case Val(Split(lblInfo(Index).Tag, ",")(3))
            Case 2
                strMask = "1234567890."
        End Select

        If InStr(strMask, Chr(KeyAscii)) = 0 And strMask <> "" Then
            KeyAscii = 0: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim strMsg As String
    
    On Error GoTo errH
    
    If txtInfo(Index).Text <> "" Then
        Select Case Val(Split(lblInfo(Index).Tag, ",")(3))
            Case 1 '日期
                 txtInfo(Index).Text = Format(zlStr.FullDate(txtInfo(Index).Text), "yyyy-MM-dd HH:mm")
                 If Not IsDate(txtInfo(Index).Text) Then
                     strMsg = "时间格式不正确,请重新录入。"
                 End If
            Case 2 '数字
                 If Not IsNumeric(txtInfo(Index).Text) Then
                     strMsg = "数字格式不正确,请重新录入。"
                 End If
        End Select
        
        If strMsg <> "" Then
            Set mobjCtl = txtInfo(Index)
            MsgBox strMsg, vbInformation, Me.Caption
            zlcontrol.TxtSelAll txtInfo(Index)
            zlcontrol.ControlSetFocus txtInfo(Index)
            Cancel = True
            Exit Sub
        End If
    End If
    
    '清除换行
    If InStr(txtInfo(Index).Text, vbCrLf) > 0 Then txtInfo(Index).Text = Replace(txtInfo(Index).Text, vbCrLf, "")

    If txtInfo(Index).Tag <> txtInfo(Index).Text And txtInfo(Index).Enabled = True And txtInfo(Index).Locked = False Then
        Call MakeText
    End If
    txtInfo(Index).Tag = txtInfo(Index).Text
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPatiInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errH
    
    If InStr("&'<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        Select Case Index
            Case idx姓名
                If txtPatiInfo(Index).Visible And txtPatiInfo(Index).Locked = False And txtPatiInfo(Index).Text <> txtPatiInfo(Index).Tag And txtPatiInfo(Index).Text <> "" Then
                    Call GetPatiList(0)
                Else
                     KeyAscii = 0
                     Call SeekNextCtl
                End If
            Case Else
                KeyAscii = 0
                Call SeekNextCtl
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPatiInfo_GotFocus(Index As Integer)
    If mblnEnter Then Call ShowCtl(txtPatiInfo(Index)): mblnEnter = False
    zlcontrol.TxtSelAll txtPatiInfo(Index)
End Sub

Private Sub optInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SeekNextCtl
    End If
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SeekNextCtl
    End If
End Sub

Private Function SeekNextCtl() As Boolean
'功能：定位到下一个焦点的控件上
    Call zlCommFun.PressKey(vbKeyTab)
    mblnEnter = True
    SeekNextCtl = True
End Function

Private Sub Subclass_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    '自定义的消息处理函数
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '鼠标坐标
    Dim intShift As Integer              '鼠标按键
    Dim bWay As Boolean                  '鼠标方向
    Dim bMouseFlag As Boolean            '鼠标事件激活标志
    Dim wzDelta, wKeys As Integer
    Select Case Msg
        Case WM_MOUSEWHEEL   '滚动
            wzDelta = (wParam And &HFFFF0000) \ &H10000 '取出32位值的高16位
            If wzDelta > 0 Then
                vscBar.Value = IIf(vscBar.Value - vscBar.LargeChange < 0, 0, vscBar.Value - vscBar.LargeChange)
            Else
                vscBar.Value = IIf(vscBar.Value + vscBar.LargeChange > vscBar.Max, vscBar.Max, vscBar.Value + vscBar.LargeChange)
            End If
    End Select
End Sub

Private Sub picSplitX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglNew As Single
    
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    If picSplitX.Tag <> "Draging" Then
        picSplitX.Tag = "Draging"
        picSplitX.BackColor = 0
    End If
    
    sglNew = picSplitX.Left + X
    
    picSplitX.Left = sglNew
End Sub

Private Sub picSplitX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    If picSplitX.Tag = "Draging" Then
        Call Form_Resize
        picSplitX.BackColor = Me.BackColor
        picSplitX.Tag = ""
    End If
End Sub

Private Sub SetPicEnabled(ByVal blnEnabled As Boolean)
    '设置控件是否可用背景色
    Dim obj As Object
    
    picEdit.Enabled = blnEnabled
    For Each obj In txtInfo
        If obj.Index <> 0 And obj.Locked = False Then
            obj.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
        End If
    Next
    
    picMainBack.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    
    '处理描述部分
    fraInfo.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    picSplitX.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    Me.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    rtbBox.Locked = Not blnEnabled
    rtbBox.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
End Sub


Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, _
    ByVal lngDeptID As Long, ByVal lngDataID As Long, ByVal lng记录ID As Long, dtBegin As Date, dtEnd As Date, btnNoEdit As Boolean, ByVal str病人类型 As String)
    On Error GoTo errH
    '参数缓存
    mInfo.EditType = 0
    mInfo.病人ID = lngPatiID
    mInfo.主页ID = lngPageID
    mInfo.交班科室ID = lngDeptID
    mInfo.内容ID = lngDataID
    mInfo.交班记录ID = lng记录ID
    mInfo.交班开始时间 = dtBegin
    mInfo.交班结束时间 = dtEnd
    mInfo.病人类型 = str病人类型
    
    '临时变量清除
    mbtnNoEdit = btnNoEdit
    mblnSave = False
    mblnChange = False
    mblnEnter = False
    Set mobjCtl = Nothing
    
    '重新加载控件
    picEdit.Enabled = True
    Call ClearData
    Call UnloadCtl
    
    '自动填充数据
    If mInfo.内容ID <> 0 Then Call LoadData
    
    Call InitCtl
    Call CtlEnabled
    Call SetPicEnabled(False)
    Call DrawLine '重新刷新界面
    
    
    If mInfo.内容ID <> 0 Then '预览

        '未保存则自动提取数据
        If rtbBox.Text = "" Then
            Call IntData
            Call MakeText
        Else
            '读取已保存的数据
           Call ReadData
        End If
    Else
        Call MakeText
    End If
    
    cmdType.Visible = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadData()
    '初始化提取数据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    
    On Error GoTo errH
    If mInfo.内容ID = 0 Then Exit Sub
    '提取已保存的数据
    strSQL = "select 内容Id, 记录id, 序号, 病人类型, 病人id, 主页id, 姓名, 性别, 年龄, 床号, 标识号, 入院时间, 入院方式, 交班描述 from 医生交接班内容 where 内容id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.内容ID)
    
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            mInfo.内容序号 = Val(rsTmp!序号 & "")
            mInfo.病人ID = Val(rsTmp!病人ID & "")
            mInfo.主页ID = Val(rsTmp!主页ID & "")
            
            txtPatiInfo(idx姓名).Text = rsTmp!姓名 & ""
            txtPatiInfo(idx性别).Text = rsTmp!性别 & ""
            txtPatiInfo(idx年龄).Text = rsTmp!年龄 & ""
            txtPatiInfo(idx床号).Text = rsTmp!床号 & ""
            txtPatiInfo(idx住院号).Text = rsTmp!标识号 & ""
            txtPatiInfo(idx入院途径).Text = rsTmp!入院方式 & ""
            txtPatiInfo(idx入院时间).Text = Format(rsTmp!入院时间 & "", "yyyy-MM-dd HH:mm")
            
            txtPatiInfo(idx病人类型).Text = rsTmp!病人类型 & ""
            mInfo.病人类型 = rsTmp!病人类型 & ""

            rtbBox.Text = rsTmp!交班描述 & ""
            rtbBox.Tag = rsTmp!交班描述 & ""
            
            With rtbBox
                .SelStart = 0
                .SelLength = Len(rtbBox.Text)
                .SelColor = vbBlack
                .SelStart = Len(rtbBox.Text)
            End With
            
            mblnSave = rsTmp!交班描述 & "" <> ""
 
            '未保存时支持修改病人类型
            If mInfo.EditType = 2 And rsTmp!交班描述 & "" = "" Then cmdType.Visible = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearData()
'功能：清空界面数据
    Dim obj As Object

    On Error GoTo errH
    For Each obj In txtPatiInfo
        obj.Text = ""
        obj.Tag = ""
    Next
    
    rtbBox.Text = ""
    rtbBox.Tag = ""
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub GetPatiList(ByVal intType As Integer)
'功能：获取当前科室的病人
'参数：0 文本框按回车，1 点按钮
    Dim strSQL As String, rsTmp As Recordset, rsTmp1 As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtPatiInfo(idx姓名).Tag = txtPatiInfo(idx姓名).Text And txtPatiInfo(idx姓名).Text <> "" Then
            Call SeekNextCtl
            Exit Sub
        End If
        
        '录入项为空时，加载全部
        If txtPatiInfo(idx姓名).Text = "" Then intType = 1
    End If
            
    strInput = Trim(UCase(txtPatiInfo(idx姓名).Text))   '传入的值存在前缀空格

    strSQL = "Select b.病人ID as ID,A.主页ID,b.姓名, b.性别, b.年龄, b.当前床号 As 床号, a.住院号, a.入院方式, a.入院日期" & vbNewLine & _
            "From 病案主页 A, 病人信息 B, 在院病人 C" & vbNewLine & _
            "Where c.病人id = a.病人id And c.主页id = a.主页id And a.病人id = b.病人id And C.科室id = [1] And a.病人性质 In(0,2)" & _
            IIf(intType = 0, " And (A.住院号 = [2] Or A.姓名 Like [3] or b.当前床号 like [3])", "") & " ORDER BY a.入院日期 desc"
        
    vRect = zlcontrol.GetControlRect(txtPatiInfo(idx姓名).hWnd)
    Set mobjCtl = txtPatiInfo(idx姓名)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "本科室病人列表", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtPatiInfo(idx姓名).Height, blnCancel, False, True, mInfo.交班科室ID, Val(strInput), mstrLike & strInput & "%", CDate(mInfo.交班开始时间), CDate(mInfo.交班结束时间))
    
    zlcontrol.ControlSetFocus txtPatiInfo(idx姓名)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Set mobjCtl = txtPatiInfo(idx姓名)
            Call MsgBox("没有在当前交班科室找到匹配的病人!", vbInformation, gstrSysName)
        End If
        txtPatiInfo(idx姓名).Text = txtPatiInfo(idx姓名).Tag
        blnDo = False
    Else
        If Not rsTmp.EOF Then
            blnDo = True
        Else
            txtPatiInfo(idx姓名).Text = txtPatiInfo(idx姓名).Tag
            blnDo = False
        End If
    End If
    
    If blnDo Then

        '判断当前病人是否存在
        strSQL = "select 内容Id,病人类型,姓名 from 医生交接班内容 where 记录id=[1] and 病人ID=[2] AND 主页ID=[3]"
        Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.交班记录ID, Val(rsTmp!id & ""), Val(rsTmp!主页ID & ""))

        If Not rsTmp1 Is Nothing Then
            If Not rsTmp1.EOF Then
                 Set mobjCtl = txtPatiInfo(idx姓名)
                 Call MsgBox("在当前交班记录中已存在当前选择的病人,请重新选择!", vbInformation, gstrSysName)
                 zlcontrol.ControlSetFocus txtPatiInfo(idx姓名)
                 txtPatiInfo(idx姓名).Text = txtPatiInfo(idx姓名).Tag
                 Call zlcontrol.TxtSelAll(txtPatiInfo(idx姓名))
                 Exit Sub
            End If
        End If

        '清除界面数据
        Call ClearData

        '加载病人信息
        mInfo.病人ID = rsTmp!id & ""
        mInfo.主页ID = rsTmp!主页ID & ""
        mInfo.病人类型 = GetPatiType(Val(rsTmp!id & ""), Val(rsTmp!主页ID & ""))
        txtPatiInfo(idx病人类型).Text = mInfo.病人类型
        txtPatiInfo(idx姓名).Text = rsTmp!姓名 & ""
        txtPatiInfo(idx姓名).Tag = rsTmp!姓名 & ""
        txtPatiInfo(idx性别).Text = rsTmp!性别 & ""
        txtPatiInfo(idx年龄).Text = rsTmp!年龄 & ""
        txtPatiInfo(idx床号).Text = rsTmp!床号 & ""
        txtPatiInfo(idx住院号).Text = rsTmp!住院号 & ""
        txtPatiInfo(idx入院途径).Text = rsTmp!入院方式 & ""
        txtPatiInfo(idx入院时间).Text = Format(rsTmp!入院日期 & "", "yyyy-MM-dd HH:mm")
        
        Call UnloadCtl
        Call InitCtl
        Call IntData '自动填充数据
        Call MakeText

        zlcontrol.ControlSetFocus txtPatiInfo(idx姓名)
        Call SeekNextCtl
    Else
        zlcontrol.ControlSetFocus txtPatiInfo(idx姓名)
        Call zlcontrol.TxtSelAll(txtPatiInfo(idx姓名))
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim str类型 As String
    Dim objControl As CommandBarControl
    On Error GoTo errH
    Select Case Control.id
        Case ID_病案
            If Not gobjPublicAdvice Is Nothing Then
                If mInfo.病人ID <> 0 And mInfo.主页ID <> 0 Then
                    Call gobjPublicAdvice.ShowArchive(Me, mInfo.病人ID, mInfo.主页ID)
                End If
            End If
        Case ID_保存
            '项目文本框未触发MakeText时
            If Me.ActiveControl.Name = "txtInfo" Then Call MakeText
            
            If CheckData Then
                Call SaveData
                Call LoadData
                Call ReadData
                '预览
                mInfo.EditType = 0
                Call SetPicEnabled(False)
                If Not gfrmParent Is Nothing Then
                    Call gfrmParent.SetEnable
                    Call gfrmParent.RefreshEdit(mInfo.内容ID)
                End If
                mblnChange = False
                cmdType.Visible = False
            End If
        Case ID_取消
            If mblnChange And mInfo.病人ID <> 0 Then
                If MsgBox("当前界面数据已发生改变，请确认是否取消编辑？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            Call SetPicEnabled(False)
            If Not gfrmParent Is Nothing Then
                Call gfrmParent.SetEnable
                If mInfo.EditType = 1 Then
                    Call gfrmParent.RefreshEdit(mInfo.内容ID)
                ElseIf mInfo.EditType = 2 And mblnChange Then
                    Call zlRefresh(mInfo.病人ID, mInfo.主页ID, mInfo.交班科室ID, mInfo.内容ID, mInfo.交班记录ID, mInfo.交班开始时间, mInfo.交班结束时间, mbtnNoEdit, mInfo.病人类型)
                End If
            End If
            mInfo.EditType = 0
            
            Call CtlEnabled
            cmdType.Visible = False
            
            mblnChange = False
        Case ID_新增
            If mbtnNoEdit Then Exit Sub
            If mInfo.交班记录ID = 0 Then Exit Sub
            If Not gfrmParent Is Nothing Then Call gfrmParent.SetEnable(1)
            Call SetPicEnabled(True)
            cmdType.Visible = True
            Call ClearData
            mInfo.EditType = 1: mInfo.病人ID = 0: mInfo.主页ID = 0: mInfo.内容ID = 0:  mInfo.内容序号 = 0: mInfo.病人类型 = ""
            '设置控件状态
            Call UnloadCtl
            Call InitCtl
            Call CtlEnabled
            
            mblnChange = False
            cmdType.Visible = True
            Call MakeText
        Case ID_修改
            Call EditState
            cmdType.Visible = Not mblnSave
            mblnChange = False
        Case ID_删除
            If mbtnNoEdit Then Exit Sub
            If mInfo.交班记录ID = 0 Or mInfo.内容ID = 0 Then Exit Sub
            If DelEdit Then
                Call gfrmParent.RefreshEdit(mInfo.内容ID)
            End If
            mInfo.EditType = 0
            cmdType.Visible = False
            mblnChange = False
        Case Else
            If Control.id <> ID_类型 Then '预览病人类型
                Call txtPatiInfo(idx姓名).SetFocus
                
                
                For Each objControl In cbsExec(2).Controls
                    If objControl.id = Control.id Then
                        objControl.IconId = IIf(objControl.IconId = 10, 12, 10)
                    End If
                    
                    If objControl.id > ID_类型 And objControl.IconId = 12 Then
                        str类型 = str类型 & "," & objControl.Caption
                    End If
                Next


                txtPatiInfo(idx病人类型).Text = Mid(str类型, 2)
                mInfo.病人类型 = txtPatiInfo(idx病人类型).Text
                Call UnloadCtl
                Call InitCtl
                Call IntData
                Call MakeText
            End If
            
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case ID_病案
            Control.Visible = (Not mbln预览)
        Case ID_保存
            Control.Visible = (Not mbln预览) And mInfo.EditType <> 0
        Case ID_取消
            Control.Visible = (Not mbln预览) And mInfo.EditType <> 0
        Case ID_新增
            Control.Visible = (Not mbln预览) And mInfo.EditType = 0 And (Not mbtnNoEdit)
        Case ID_修改
            Control.Visible = (Not mbln预览) And mInfo.EditType = 0 And mInfo.内容ID <> 0 And (Not mbtnNoEdit)
        Case ID_删除
            Control.Visible = (Not mbln预览) And mInfo.EditType = 0 And mInfo.内容ID <> 0 And (Not mbtnNoEdit)
        Case Else
            Control.Visible = mbln预览
    End Select
End Sub


Public Sub EditState()
'功能：修改状态
    Dim i As Long
    On Error GoTo errH
    If mInfo.交班记录ID = 0 Or mInfo.内容ID = 0 Or mbtnNoEdit Then Exit Sub
    If Not gfrmParent Is Nothing Then Call gfrmParent.SetEnable(1)
    Call SetPicEnabled(True)
    '处理单选控件激活跳转的问题
    mblnLoad = True
    
    For i = 1 To txtInfo.Count - 1
        If txtInfo(i).Enabled And txtInfo(i).Locked = False Then
            txtInfo(i).SetFocus
            Exit For
        End If
    Next
    
    mblnLoad = False
    mInfo.EditType = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function DelEdit() As Boolean
    Dim strSQL As String
    Dim blnTran As Boolean
    
    On Error GoTo errH
    If MsgBox("确定要删除选中的病人交班记录吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    strSQL = "Zl_医生交接班内容_Edit(2," & mInfo.内容ID & ")"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    DelEdit = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub CtlEnabled()
    On Error GoTo errH

    '显示选择病人按钮
    txtPatiInfo(idx姓名).Locked = mInfo.EditType <> 1 Or mbln预览 = True
    txtPatiInfo(idx姓名).TabStop = mInfo.EditType = 1 And mbln预览 = False
    txtPatiInfo(idx姓名).BackColor = IIf(mInfo.EditType = 1 And mbln预览 = False, &H80000005, &H8000000F)
    cmdFind.Visible = mInfo.EditType = 1 And mbln预览 = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiType(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    '获取病人类别
    Dim strSQL As String
    Dim strType As String
    Dim strSqlTmp As String, lngCount As Long, i As Long, lngBegin As Long, lngEnd As Long, lng起始 As Long, strReplace As String
    Dim rsPatiType As ADODB.Recordset
    
    On Error GoTo errH
    If mrsCtlType Is Nothing Then Exit Function
    mrsCtlType.Filter = ""
    mrsCtlType.MoveFirst
    If mrsCtlType.EOF Then Exit Function
    If lng病人ID = 0 Or mInfo.交班记录ID = 0 Then Exit Function

    '解析当前交班病人类型的SQL
    Do While Not mrsCtlType.EOF
        If mrsCtlType!提取SQL & "" <> "" Then
        
            '替换病人ID
            strSqlTmp = Replace(mrsCtlType!提取SQL & "", "病人id", "病人ID")
            strSqlTmp = Replace(strSqlTmp, "病人iD", "病人ID")
            strSqlTmp = Replace(strSqlTmp, "病人Id", "病人ID")
            
            '获取病人ID出现次数
            lngCount = (Len(strSqlTmp) - Len(Replace(strSqlTmp, "病人ID", ""))) / Len("病人ID")
            
            lngBegin = 0: lngEnd = 0: strReplace = "": lng起始 = 1 '从1开始
            '循环解析
            For i = 1 To lngCount
                lng起始 = InStr(lng起始, strSqlTmp, "病人ID") + 1
                
                lngBegin = lng起始 - 1
                lngEnd = InStr(lngBegin, strSqlTmp, "-1")

                If lngBegin + 4 < lngEnd And lngBegin <> 0 And lngEnd <> 0 Then
                    strReplace = Replace(Mid(strSqlTmp, lngBegin + 4, lngEnd - (lngBegin + 4)), " ", "")
                    strReplace = Replace(strReplace, vbCrLf, "")
                    If strReplace = "<>" Then
                        strSqlTmp = Replace(strSqlTmp, Mid(strSqlTmp, lngBegin, lngEnd + 2 - lngBegin), "病人ID=[4]")
                        lng起始 = lngEnd + 2
                    End If
                End If
            Next

        
            '替换主页ID
            strSqlTmp = Replace(strSqlTmp, "主页id", "主页ID")
            strSqlTmp = Replace(strSqlTmp, "主页iD", "主页ID")
            strSqlTmp = Replace(strSqlTmp, "主页Id", "主页ID")
            
            '获取主页ID出现次数
            lngCount = (Len(strSqlTmp) - Len(Replace(strSqlTmp, "主页ID", ""))) / Len("主页ID")
            
            lngBegin = 0: lngEnd = 0: strReplace = "": lng起始 = 1 '从1开始
            '循环解析
            For i = 1 To lngCount
                If i = 1 Then lng起始 = 1 '从1开始
                
                lng起始 = InStr(lng起始, strSqlTmp, "主页ID") + 1
                
                lngBegin = lng起始 - 1
                lngEnd = InStr(lngBegin, strSqlTmp, "-1")

                If lngBegin + 4 < lngEnd And lngBegin <> 0 And lngEnd <> 0 Then
                    strReplace = Replace(Mid(strSqlTmp, lngBegin + 4, lngEnd - (lngBegin + 4)), " ", "")
                    strReplace = Replace(strReplace, vbCrLf, "")
                    If strReplace = "<>" Then
                        strSqlTmp = Replace(strSqlTmp, Mid(strSqlTmp, lngBegin, lngEnd + 2 - lngBegin), "主页ID=[5]")
                        lng起始 = lngEnd + 2
                    End If
                End If
            Next
            
            strSQL = strSQL & " Union All " & strSqlTmp
        End If
        mrsCtlType.MoveNext
    Loop
    
    If strSQL <> "" Then
        strSQL = Mid(strSQL, 12)
        strSQL = Replace(strSQL, "[开始时间]", "[1]")
        strSQL = Replace(strSQL, "[结束时间]", "[2]")
        strSQL = Replace(strSQL, "[科室ID]", "[3]")
        
        '容错处理
        strSQL = Replace(strSQL, "病人ID<>-1", "病人ID=[4]")
        strSQL = Replace(strSQL, "病人ID <> -1", "病人ID=[4]")
        strSQL = Replace(strSQL, "主页ID<>-1", "主页ID=[5]")
        strSQL = Replace(strSQL, "主页ID <> -1", "主页ID=[5]")
        Set rsPatiType = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mInfo.交班开始时间), CDate(mInfo.交班结束时间), mInfo.交班科室ID & "", lng病人ID, lng主页ID)
        rsPatiType.Filter = "病人ID =" & lng病人ID & " And 主页ID =" & lng主页ID
    End If
    
    If Not rsPatiType Is Nothing Then
        Do While Not rsPatiType.EOF
            If InStr("," & strType & ",", "," & rsPatiType!类型 & ",") = 0 And rsPatiType!类型 & "" <> "" Then
                strType = strType & "," & rsPatiType!类型
            End If
            rsPatiType.MoveNext
        Loop
        strType = Mid(strType, 2)
    End If
    
    GetPatiType = strType
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function





Private Function ReadData() As Boolean
    '获取已保存的交班内容
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim obj As Object
    Dim strValue As String, strTmp As String
    Dim lngBegin As Long, lngEnd As Long, j As Long

    
    On Error GoTo errH
    If mInfo.内容ID = 0 Then Exit Function
    
    strSQL = "Select 内容id, 序号, 项目, 内容 From 医生交接班详情 Where 内容id = [1] Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.内容ID)
    
    If rsTmp.EOF Then Exit Function
    
    mblnLoad = True
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            rsTmp.Filter = "项目 ='" & Replace(obj.Caption, vbCrLf, "") & "'"
            
            If Not rsTmp.EOF Then
                 strValue = rsTmp!内容 & ""
                 
                 '存在选项框
                 If InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") > 0 And strValue <> "" Then
                    
                    If InStr(strValue, ";") > 0 Then
                        strTmp = Split(strValue, ";")(0)
                        strValue = Mid(strValue, Len(strTmp) + 3)
                    End If
                    
                    lngBegin = Val(Split(picTmp(obj.Index).Tag, ",")(1))
                    lngEnd = Val(Split(picTmp(obj.Index).Tag, ",")(2))
                 
                    If Val(Mid(picTmp(obj.Index).Tag, 1, 1)) = 2 Then
                        For j = lngBegin To lngEnd
                            If InStr("," & strTmp & ",", "," & optInfo(j).Caption & ",") > 0 Then
                                optInfo(j).Value = True
                            End If
                        Next
                    Else
                        For j = lngBegin To lngEnd
                            If InStr("," & strTmp & ",", "," & chkInfo(j).Caption & ",") > 0 Then
                                chkInfo(j).Value = 1
                            End If
                        Next
                    End If
                    
                    Call SetTextVisible(obj.Index)
                 End If
                 
                 '加载文本框
                 If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 And strValue <> "" Then
                    txtInfo(obj.Index) = strValue
                    txtInfo(obj.Index).Tag = strValue
                 End If
            End If
        End If
    Next
    
    mblnLoad = False
    ReadData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckData()
    '保存前检查
    Dim obj As Object
    Dim blnCheck As Boolean
    Dim lngBegin As Long, lngEnd As Long, j As Long
    
    On Error GoTo errH
    
    If mInfo.病人ID = 0 Then
        Set mobjCtl = txtPatiInfo(idx姓名)
        Call MsgBox("请选择需要填写交班记录的病人!", vbInformation, gstrSysName)
        zlcontrol.ControlSetFocus txtPatiInfo(idx姓名)
        Call zlcontrol.TxtSelAll(txtPatiInfo(idx姓名))
        Exit Function
    End If
    
    If txtPatiInfo(idx病人类型).Text = "" Then
        Set mobjCtl = txtPatiInfo(idx病人类型)
        Call MsgBox("请选择当前病人的病人类型!", vbInformation, gstrSysName)
        zlcontrol.ControlSetFocus txtPatiInfo(idx病人类型)
        Call zlcontrol.TxtSelAll(txtPatiInfo(idx病人类型))
        Exit Function
    End If
    
    
    For Each obj In lblInfo
        If obj.Index <> 0 Then
                '存在选项框
                If InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") > 0 Then
                   lngBegin = Val(Split(picTmp(obj.Index).Tag, ",")(1))
                   lngEnd = Val(Split(picTmp(obj.Index).Tag, ",")(2))
                    blnCheck = False
                   If Val(Mid(picTmp(obj.Index).Tag, 1, 1)) = 3 And lngEnd <> 0 And lngBegin <> lngEnd Then
                       For j = lngBegin To lngEnd
                           If chkInfo(j).Value = 1 Then
                               blnCheck = True
                               Exit For
                           End If
                       Next
                       
                        If blnCheck = False Then
                            If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") = 0 Then
                                 Call ShowMsg(obj.Caption & "没有选择任何选项,不能保存!", chkInfo(lngBegin))
                                 Exit Function
                            Else
                                If txtInfo(obj.Index) = "" Then
                                    Call ShowMsg(obj.Caption & "没有选择任何选项,不能保存!", chkInfo(lngBegin))
                                    Exit Function
                                End If
                            End If
                        End If
                   End If
                End If
                
                '检查文本框是否为空
                If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 Then
                    If txtInfo(obj.Index).Text = "" And txtInfo(obj.Index).Locked = False And txtInfo(obj.Index).Visible = True And InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") = 0 Then
                        Call ShowMsg(obj.Caption & "不能为空，请重新录入!", txtInfo(obj.Index))
                        Exit Function
                    End If
                    
                    '检查输入类型是否正确
                    Select Case Val(Split(obj.Tag, ",")(3))
                            Case 1
                                If (Not IsDate(txtInfo(obj.Index).Text)) And txtInfo(obj.Index).Locked = False And txtInfo(obj.Index).Visible = True Then
                                    Call ShowMsg(obj.Caption & "不是有效的日期，请重新录入!", txtInfo(obj.Index))
                                    Exit Function
                                End If
                            Case 2
                                If (Not IsNumeric(txtInfo(obj.Index).Text)) And txtInfo(obj.Index).Locked = False And txtInfo(obj.Index).Visible = True Then
                                    Call ShowMsg(obj.Caption & "不是有效的数字，请重新录入!", txtInfo(obj.Index))
                                    Exit Function
                                End If
                    End Select
                End If
        End If
    Next
    
    CheckData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Function ShowMsg(strMsg As String, obj As Object)
    '用于检查提示时定位控件
    On Error Resume Next
    '信息提示
    Set mobjCtl = obj
    Call MsgBox(strMsg, vbInformation, gstrSysName)
    zlcontrol.ControlSetFocus obj
    If obj.Name = "txtInfo" Then
        Call zlcontrol.TxtSelAll(txtInfo(obj.Index))
    End If

    '控件定位
    Call ShowCtl(obj)
End Function


Private Function ShowCtl(obj As Object)
    '用于检查提示时定位控件
    Dim lngTop1 As Long, lngTop2 As Long

    On Error Resume Next
    '控件定位
    If vscBar.Visible Then
        Select Case obj.Name
                Case "txtInfo"
                    lngTop1 = obj.Top + obj.Height + 100
                    lngTop2 = obj.Top
                Case "chkInfo", "optInfo"
                    lngTop1 = obj.Container.Top + obj.Container.Height + 100
                    lngTop2 = obj.Container.Top
        End Select

        If lngTop1 < picMain.Height Then
            vscBar.Value = vscBar.Min
        ElseIf picEdit.Height - lngTop2 < picMain.Height Then
            vscBar.Value = vscBar.Max
        Else
            vscBar.Value = lngTop1 - picMain.Height
        End If
        zlcontrol.ControlSetFocus obj
    End If
End Function


Private Function GetDateStr(ByVal intIndex As Integer) As String
    '获取日期格式字符串
    Dim strType As String, strName As String
    Dim strValue As String
    On Error GoTo errH
    If intIndex = 0 Then Exit Function
    
    strType = Split(lblInfo(intIndex).Tag, ",")(1)
    strName = Replace(lblInfo(intIndex).Caption, vbCrLf, "")

    If Not mrsCtlInfo Is Nothing Then
        mrsCtlInfo.Filter = "病人简称 ='" & strType & "' And 项目名称 ='" & strName & "'"
        If Not mrsCtlInfo.EOF Then strValue = mrsCtlInfo!输入格式 & ""
        mrsCtlInfo.Filter = ""
    End If
    
    If strValue = "" Then strValue = "yyyy-MM-dd HH:mm" '默认格式
    
    GetDateStr = strValue
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitRecordset(rsTmp As Recordset)
'功能：初始化病历提取记录集
    Set rsTmp = New ADODB.Recordset
    
    rsTmp.Fields.Append "病历类型", adVarChar, 5000
    rsTmp.Fields.Append "提纲名称", adVarChar, 5000
    rsTmp.Fields.Append "提纲取值", adVarChar, 5000
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
End Sub

Private Function IntData() As Boolean
    '初始化提取数据
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, rsValue As ADODB.Recordset, rs病历提取 As ADODB.Recordset
    Dim obj As Object
    Dim strValue As String
    Dim str诊断 As String, str体征 As String, str输血情况 As String
    Dim arrTmp As Variant, i As Long
    
    On Error GoTo errH
    If mrsCtlInfo Is Nothing Then Exit Function
    
    mrsCtlInfo.Filter = ""
    If mInfo.病人ID = 0 Or mrsCtlInfo.EOF Or mInfo.病人类型 = "" Then Exit Function

    Call InitRecordset(rs病历提取)

    '首先把自定义数据源的SQL一次读取出来
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 Then
                mrsCtlInfo.Filter = "项目名称='" & Replace(obj.Caption, vbCrLf, "") & "' And 病人简称 ='" & Split(obj.Tag, ",")(1) & "'"
                If Not mrsCtlInfo.EOF Then
                    '自定义sql拼接
                    If Val(mrsCtlInfo!输入形式 & "") = 1 And Val(mrsCtlInfo!提取来源 & "") = 99 And mrsCtlInfo!提取SQL <> "" Then
                            strSQL = strSQL & " Union All Select '" & Replace(obj.Caption, vbCrLf, "") & "' As 项目名称, a.* From (" & mrsCtlInfo!提取SQL & ") A "
                    
                    '病历提纲拼接
                    ElseIf Val(mrsCtlInfo!输入形式 & "") = 1 And Val(mrsCtlInfo!提取来源 & "") = 4 And mrsCtlInfo!提取病历 & "" <> "" And InStr(mrsCtlInfo!提取病历 & "", ":") > 0 Then
                            rs病历提取.Filter = "病历类型 ='" & Split(mrsCtlInfo!提取病历 & "", ":")(0) & "'"
                            If rs病历提取.EOF Then
                                rs病历提取.AddNew
                                rs病历提取!病历类型 = Split(mrsCtlInfo!提取病历 & "", ":")(0)
                                rs病历提取!提纲名称 = Split(mrsCtlInfo!提取病历 & "", ":")(1)
                            Else
                                arrTmp = Array()
                                arrTmp = Split(Split(mrsCtlInfo!提取病历 & "", ":")(1), ";")
                                For i = 0 To UBound(arrTmp)
                                    If arrTmp(i) <> "" And InStr(";" & rs病历提取!提纲名称 & ";", ";" & arrTmp(i) & ";") = 0 Then
                                        rs病历提取!提纲名称 = rs病历提取!提纲名称 & ";" & arrTmp(i)
                                    End If
                                Next
                            End If
                    End If
                End If
            End If
        End If
    Next
    
    '一次读取病历提取
    rs病历提取.Filter = ""
    Do While Not rs病历提取.EOF
        If rs病历提取!病历类型 & "" <> "" And rs病历提取!提纲名称 <> "" Then
            rs病历提取!提纲取值 = GetOPSByEmr(rs病历提取!病历类型 & "", mInfo.病人ID, mInfo.主页ID, Replace(rs病历提取!提纲名称 & "", ";", "|"))
            rs病历提取!提纲取值 = Replace(rs病历提取!提纲取值 & "", vbCrLf, "")
            rs病历提取!提纲取值 = Replace(rs病历提取!提纲取值 & "", Chr(10), "")
            rs病历提取!提纲取值 = Replace(rs病历提取!提纲取值 & "", Chr(13), "")
        End If
        rs病历提取.MoveNext
    Loop
    
    
    If strSQL <> "" Then
        strSQL = Mid(strSQL, 12)
        strSQL = Replace(strSQL, "[病人ID]", "[1]")
        strSQL = Replace(strSQL, "[主页ID]", "[2]")
        strSQL = Replace(strSQL, "[开始时间]", "[3]")
        strSQL = Replace(strSQL, "[结束时间]", "[4]")
        strSQL = Replace(strSQL, "[科室ID]", "[5]")
        
        Set rsValue = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.病人ID, mInfo.主页ID, CDate(mInfo.交班开始时间), CDate(mInfo.交班结束时间), mInfo.交班科室ID & "")
    End If


    For Each obj In lblInfo
        If obj.Index <> 0 Then
            '只提取文本框初始数据
            If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 Then
                mrsCtlInfo.Filter = "项目名称='" & Replace(obj.Caption, vbCrLf, "") & "' And 病人简称 ='" & Split(obj.Tag, ",")(1) & "'"
                If Not mrsCtlInfo.EOF Then
                    If Val(mrsCtlInfo!输入形式 & "") = 1 Then
                            Select Case Val(mrsCtlInfo!提取来源 & "")
                                Case 1 '提取最新诊断
                                    If str诊断 = "" Then
                                        strSQL = "Select Zl_Fun_Getpatishift(1, [1], [2], [3], [4]) as 最新诊断 From Dual"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.病人ID, mInfo.主页ID, CDate(mInfo.交班开始时间), CDate(mInfo.交班结束时间))
                                        If Not rsTmp Is Nothing Then str诊断 = rsTmp!最新诊断 & ""
                                     End If
                                    txtInfo(obj.Index).Text = str诊断
                                Case 2 '提取最新体征
                                    If str体征 = "" Then
                                        strSQL = "Select Zl_Fun_Getpatishift(2, [1], [2], [3], [4]) as 病人体征 From Dual"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.病人ID, mInfo.主页ID, CDate(mInfo.交班开始时间), CDate(mInfo.交班结束时间))
                                        If Not rsTmp Is Nothing Then str体征 = rsTmp!病人体征 & ""
                                    End If
                                    txtInfo(obj.Index).Text = str体征
                                Case 3 '提取输血情况
                                    If str输血情况 = "" Then
                                        strSQL = "Select Zl_Fun_Getpatishift(3, [1], [2], [3], [4]) as 输血情况 From Dual"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.病人ID, mInfo.主页ID, CDate(mInfo.交班开始时间), CDate(mInfo.交班结束时间))
                                        If Not rsTmp Is Nothing Then str输血情况 = rsTmp!输血情况 & ""
                                    End If
                                    txtInfo(obj.Index).Text = str输血情况
                                Case 4 '提取新版病历内容
                                    strValue = Get病历提纲(mrsCtlInfo!提取病历 & "", rs病历提取)
                                    If strValue & "" <> "" Then
                                       If Val(mrsCtlInfo!输入类型 & "") = 0 Then
                                            txtInfo(obj.Index).Text = strValue
                                       ElseIf Val(mrsCtlInfo!输入类型 & "") = 1 Then
                                            txtInfo(obj.Index).Text = Format(strValue, "yyyy-MM-dd HH:mm")
                                       ElseIf Val(mrsCtlInfo!输入类型 & "") = 2 Then
                                            txtInfo(obj.Index).Text = Val(strValue)
                                       End If
                                    End If
                                Case 99 '通过SQL提取
                                    If Not rsValue Is Nothing Then
                                        rsValue.Filter = "项目名称 = '" & Replace(obj.Caption, vbCrLf, "") & "'"
                                        If Not rsValue.EOF Then
                                            If rsValue.Fields(1).Value & "" <> "" Then
                                               If Val(mrsCtlInfo!输入类型 & "") = 0 Then
                                                    txtInfo(obj.Index).Text = rsValue.Fields(1).Value & ""
                                               ElseIf Val(mrsCtlInfo!输入类型 & "") = 1 Then
                                                    txtInfo(obj.Index).Text = Format(rsValue.Fields(1).Value & "", "yyyy-MM-dd HH:mm")
                                               ElseIf Val(mrsCtlInfo!输入类型 & "") = 2 Then
                                                    txtInfo(obj.Index).Text = Val(rsValue.Fields(1).Value & "")
                                               End If
                                            End If
                                        End If
                                    End If
                            End Select
                    End If
                End If
                txtInfo(obj.Index).Tag = txtInfo(obj.Index).Text
            End If
        End If
    Next
    
    mblnChange = False
    
    IntData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get病历提纲(ByVal str格式 As String, rsTmp As ADODB.Recordset) As String
    '提取新版病历提纲
    Dim arrTmp As Variant, arr类型 As Variant, arr取值 As Variant, i As Long, j As Long
    Dim strTmp As String
    Dim strValue As String
    
    On Error GoTo errH
    
    '容错处理
    If InStr(str格式, ":") = 0 Or mInfo.病人ID = 0 Then Exit Function
    If Split(str格式, ":")(1) = "" Or Split(str格式, ":")(0) = "" Then Exit Function
    If rsTmp Is Nothing Then Exit Function
    rsTmp.Filter = "病历类型 ='" & Split(str格式, ":")(0) & "'"
    If rsTmp.EOF Then Exit Function
    If rsTmp!提纲取值 & "" = "" Then Exit Function
    arrTmp = Array()
    arrTmp = Split(Split(str格式, ":")(1), ";")
    arr类型 = Array()
    arr类型 = Split(rsTmp!提纲名称 & "", ";")
    arr取值 = Array()
    arr取值 = Split(rsTmp!提纲取值 & "", "|")
    
    If UBound(arr类型) <> UBound(arr取值) Then Exit Function

    For i = 0 To UBound(arrTmp)
        For j = 0 To UBound(arr类型)
            If arrTmp(i) = arr类型(j) Then
                strValue = arr取值(j)
                strValue = Replace(strValue, arr类型(j) & " : ", "")
                strValue = Replace(strValue, arr类型(j) & ": ", "")
                strValue = Replace(strValue, arr类型(j) & ":", "")
                strValue = Replace(strValue, arr类型(j) & " :", "")
                
                strValue = Replace(strValue, arr类型(j) & " ： ", "")
                strValue = Replace(strValue, arr类型(j) & "： ", "")
                strValue = Replace(strValue, arr类型(j) & "：", "")
                strValue = Replace(strValue, arr类型(j) & " ：", "")

                strTmp = strTmp & ";" & strValue
                Exit For
            End If
        Next
    Next
    
    strTmp = Mid(strTmp, 2)

    Get病历提纲 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetOPSByEmr(ByVal strDocKind As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str提纲 As String) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等。从病历中获取附项值
    Dim strText As String
    
    On Error Resume Next
    
    If gobjEmr Is Nothing Then Exit Function
    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing: Exit Function
 
    If Not gobjEmr Is Nothing Then
        strText = gobjEmr.GetContentOfSpecifyDoc(strDocKind, lng病人ID, lng主页ID, str提纲)
    End If
    
    Err.Clear
    GetOPSByEmr = strText
End Function

Private Sub MakeText()
     '生成交班描述
    Dim str描述 As String, strS描述 As String, str主诉 As String, str诊断 As String, strB描述 As String, strA描述 As String, strR描述 As String
    Dim strType As String, strTmp As String, strValue As String
    Dim obj As Object, i As Long, j As Long, lngBegin As Long, lngEnd As Long
    Dim dtNow As Date, dtTmp As Date
        
    On Error GoTo errH
    
    dtNow = zlDatabase.Currentdate
    
    With rtbBox
        .SelStart = 0
        .SelLength = Len(rtbBox.Text)
        .SelColor = vbBlack
        .SelStart = Len(rtbBox.Text)
    End With
    
    If mInfo.病人ID = 0 Then
        str描述 = "[床号]患者[姓名]，[性别]，[年龄]，住院号：[住院号]，[入院时间]以[主诉]为主诉[入院途径]入院。"
        rtbBox.Text = str描述
        rtbBox.Tag = str描述
        Exit Sub
    End If
    '生成病人信息和S类项目的描述
    str描述 = "[床号]患者[姓名][性别][年龄]，住院号：[住院号][strS描述]，[入院时间][str主诉][入院途径]入院[str诊断]。[strB描述][strA描述][strR描述]"
    
    For Each obj In lblPatiInfo
        If obj.Caption = "性别" Or obj.Caption = "年龄" Then
            str描述 = Replace(str描述, "[" & obj.Caption & "]", IIf(txtPatiInfo(obj.Index).Text = "", "", "，" & txtPatiInfo(obj.Index).Text))
        ElseIf obj.Caption = "床号" Then
            str描述 = Replace(str描述, "[" & obj.Caption & "]", IIf(txtPatiInfo(obj.Index).Text = "", "", txtPatiInfo(obj.Index).Text & "床"))
        ElseIf obj.Caption = "入院时间" And IsDate(txtPatiInfo(obj.Index).Text) And txtPatiInfo(obj.Index).Text <> "" Then
            dtTmp = CDate(txtPatiInfo(obj.Index).Text)
            If Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) Then
                strTmp = "今日" & Format(txtPatiInfo(obj.Index).Text, "HH时mm分")
            ElseIf Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) + 1 Then
                strTmp = "明日" & Format(txtPatiInfo(obj.Index).Text, "HH时mm分")
            Else
                strTmp = IIf(txtPatiInfo(obj.Index).Text = "", "", Format(txtPatiInfo(obj.Index).Text, "yyyy年MM月dd日HH时mm分"))
            End If
            str描述 = Replace(str描述, "[" & obj.Caption & "]", strTmp)
        Else
            str描述 = Replace(str描述, "[" & obj.Caption & "]", txtPatiInfo(obj.Index).Text)
        End If
    Next
    

    For i = 1 To lblInfo.Count - 1
        If lblInfo(i).Caption = "主诉" Then '生成主诉的描述
            str主诉 = IIf(txtInfo(i).Text = "", "", "以“" & txtInfo(i).Text & "”为主诉")
        ElseIf InStr(lblInfo(i).Caption, "诊断") > 0 Then '生成诊断的描述
            If txtInfo(i).Text <> "" Then
                str诊断 = "，" & lblInfo(i).Caption & "：" & txtInfo(i).Text
            End If
        Else
            mrsCtlInfo.Filter = "项目名称='" & Replace(lblInfo(i).Caption, vbCrLf, "") & "' And 病人简称 ='" & Split(lblInfo(i).Tag, ",")(1) & "'"
            If Not mrsCtlInfo.EOF Then
                strValue = "": strTmp = ""
                If mrsCtlInfo!描述文字 & "" = "" Then
                    strValue = lblInfo(i).Caption & "："
                Else
                    If mrsCtlInfo!描述文字 & "" = "-" Then
                        strValue = ""
                    Else
                        strValue = mrsCtlInfo!描述文字 & "："
                    End If
                End If
                
                If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
                   lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
                   lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
                   
                   If Val(Mid(picTmp(i).Tag, 1, 1)) = 3 And lngEnd <> 0 Then
                       For j = lngBegin To lngEnd
                           If chkInfo(j).Value = 1 Then
                                strTmp = strTmp & IIf(strTmp = "", "", "，") & chkInfo(j).Caption
                           End If
                       Next
                   ElseIf Val(Mid(picTmp(i).Tag, 1, 1)) = 2 And lngEnd <> 0 Then
                        For j = lngBegin To lngEnd
                           If optInfo(j).Value = True Then
                                strTmp = strTmp & IIf(strTmp = "", "", "，") & optInfo(j).Caption
                                Exit For
                           End If
                        Next
                   End If
                End If
                
                If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 Then
                   If Val(mrsCtlInfo!输入类型 & "") = 1 And Val(mrsCtlInfo!输入形式 & "") = 1 And txtInfo(i).Text <> "" And IsDate(txtInfo(i).Text) Then
                        dtTmp = CDate(txtInfo(i).Text)
                        If Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) Then
                            strTmp = strTmp & IIf(strTmp = "", "", "，") & "今日" & Format(txtInfo(i).Text, "HH时mm分")
                        ElseIf Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) + 1 Then
                            strTmp = strTmp & IIf(strTmp = "", "", "，") & "明日" & Format(txtInfo(i).Text, "HH时mm分")
                        Else
                            strTmp = strTmp & IIf(strTmp = "", "", "，") & IIf(txtInfo(i).Text = "", "", Format(txtInfo(i).Text, GetDateStr(i)))
                        End If
                    
                   ElseIf txtInfo(i).Text <> "" Then
                       strTmp = strTmp & IIf(strTmp = "", "", "，") & txtInfo(i).Text
                   End If
                End If
                
                '加载B类型的起始描述
                If Split(lblInfo(i).Tag, ",")(2) = "B" Then
                    If InStr("," & strType & ",", "," & Split(lblInfo(i).Tag, ",")(1) & ",") = 0 Then
                        mrsCtlType.Filter = "简称 ='" & Split(lblInfo(i).Tag, ",")(1) & "'"
                        
                        If Not mrsCtlType.EOF Then
                            If mrsCtlType!起始描述 & "" <> "" Then
                                strB描述 = strB描述 & IIf(Right(strB描述, 1) <> "，" And Right(strB描述, 1) <> "。" And strB描述 <> "", "，", "") & mrsCtlType!起始描述
                            End If
                        End If
                        strType = strType & "," & Split(lblInfo(i).Tag, ",")(1)
                    End If
                End If
                
                If strTmp <> "" Then
                    Select Case Split(lblInfo(i).Tag, ",")(2)
                            Case "S"
                                strS描述 = strS描述 & IIf(strS描述 = "", "", "，") & strValue & strTmp
                            Case "B"
                                If InStr(strB描述, "[" & Replace(lblInfo(i).Caption, vbCrLf, "") & "]") > 0 Then
                                    strB描述 = Replace(strB描述, "[" & Replace(lblInfo(i).Caption, vbCrLf, "") & "]", strTmp)
                                Else
                                    strB描述 = strB描述 & IIf(strB描述 = "" Or Right(strB描述 & "", 1) = "：" Or Right(strB描述 & "", 1) = "，" Or Right(strB描述 & "", 1) = "。", "", "，") & strValue & strTmp
                                End If
                            Case "A"
                                strA描述 = strA描述 & IIf(strS描述 = "", "", "，") & strValue & strTmp
                            Case "R"
                                strR描述 = strR描述 & IIf(strS描述 = "", "", "，") & strValue & strTmp
                    End Select
                End If
            End If
        End If
    Next

    str描述 = Replace(str描述, "[str主诉]", str主诉)
    str描述 = Replace(str描述, "[str诊断]", str诊断)
    str描述 = Replace(str描述, "[strS描述]", strS描述)
    str描述 = Replace(str描述, "[strA描述]", strA描述 & IIf(strA描述 = "", "", "。"))
    str描述 = Replace(str描述, "[strR描述]", strR描述 & IIf(strR描述 = "", "", "。"))
    str描述 = Replace(str描述, "[strB描述]", strB描述 & IIf(strB描述 = "", "", "。"))

    '特殊字符处理
    str描述 = Replace(str描述, "&", "")
    str描述 = Replace(str描述, "'", "")
    str描述 = Replace(str描述, "<", "")
    str描述 = Replace(str描述, ">", "")
    rtbBox.Text = str描述
    rtbBox.Tag = str描述
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function SaveData()
    '保存数据
    Dim i As Long, j As Long, lngBegin As Long, lngEnd As Long
    Dim strValue As String, blnTran As Boolean
    Dim arrSQL As Variant

    On Error GoTo errH

    arrSQL = Array()
    '更新内容记录
    If mInfo.EditType = 1 Then  '新增
        mInfo.内容ID = GetNextId("医生交接班内容", "内容ID")
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班内容_Edit(0," & mInfo.内容ID & "," & mInfo.交班记录ID & "," & mInfo.内容序号 & ",'" & mInfo.病人类型 & "'," & mInfo.病人ID & "," & mInfo.主页ID & ",'" & _
                            txtPatiInfo(idx姓名).Text & "','" & txtPatiInfo(idx性别).Text & "','" & txtPatiInfo(idx年龄).Text & "','" & txtPatiInfo(idx床号).Text & "'," & ZVal(txtPatiInfo(idx住院号).Text) & _
                            ",to_date('" & Format(txtPatiInfo(idx入院时间).Text, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ",'" & _
                            txtPatiInfo(idx入院途径).Text & "','" & rtbBox.Text & "')"
    ElseIf mInfo.EditType = 2 Then  '修改
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班内容_Edit(1," & mInfo.内容ID & "," & mInfo.交班记录ID & "," & mInfo.内容序号 & ",'" & mInfo.病人类型 & "'," & mInfo.病人ID & "," & mInfo.主页ID & ",'" & _
                            txtPatiInfo(idx姓名).Text & "','" & txtPatiInfo(idx性别).Text & "','" & txtPatiInfo(idx年龄).Text & "','" & txtPatiInfo(idx床号).Text & "'," & ZVal(txtPatiInfo(idx住院号).Text) & _
                            ",to_date('" & Format(txtPatiInfo(idx入院时间).Text, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ",'" & _
                            txtPatiInfo(idx入院途径).Text & "','" & rtbBox.Text & "')"
    End If
    
    '更新内容详情
    If mInfo.EditType = 2 Then  '修改时先删除在重新提交
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班详情_Edit(2," & mInfo.内容ID & ")"
    End If
    
    For i = 1 To lblInfo.Count - 1
        '获取控件内容
        strValue = ""
        
         '存在选项框
         If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
            
            lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
            lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
         
            If Val(Mid(picTmp(i).Tag, 1, 1)) = 2 Then
                For j = lngBegin To lngEnd
                    If optInfo(j).Value Then
                        strValue = strValue & "," & optInfo(j).Caption
                    End If
                Next
            Else
                For j = lngBegin To lngEnd
                    If chkInfo(j).Value = 1 Then
                        strValue = strValue & "," & chkInfo(j).Caption
                    End If
                Next
            End If
            strValue = Mid(strValue, 2)
            strValue = IIf(strValue <> "", strValue & ";", "")
         End If
         
         '加载文本框
         If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 Then
            strValue = strValue & txtInfo(i).Text
         End If
         
        strValue = Replace(strValue, "'", "")
        '组成SQL
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班详情_Edit(1," & mInfo.内容ID & "," & i & ",'" & Replace(lblInfo(i).Caption, vbCrLf, "") & "','" & strValue & "')"
    Next
    
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Debug.Print CStr(arrSQL(i))
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    
    If mInfo.EditType = 1 Then  '同步记录序号
        mInfo.内容序号 = Val(Sys.RowValue("医生交接班内容", mInfo.内容ID, "序号", "内容ID") & "")
    End If
    
    On Error GoTo 0
    Screen.MousePointer = 0
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function
