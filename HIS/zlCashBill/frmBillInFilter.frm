VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillInFilter 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "过滤条件"
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt登记人 
      Height          =   315
      Left            =   4890
      TabIndex        =   9
      Top             =   75
      Width           =   1650
   End
   Begin VB.CommandButton cmd刷新 
      Caption         =   "过滤(&F)"
      Height          =   390
      Left            =   5475
      TabIndex        =   8
      Top             =   540
      Width           =   1050
   End
   Begin VB.CheckBox chk仅有余额 
      BackColor       =   &H8000000A&
      Caption         =   "仅显示有库存数的记录(&N)"
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   630
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   3
      Left            =   615
      TabIndex        =   2
      Top             =   2250
      Width           =   3105
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   162332675
      CurrentDate     =   37007
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   162332675
      CurrentDate     =   37007
   End
   Begin VB.Label lbl登记人 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登记人"
      Height          =   180
      Left            =   4335
      TabIndex        =   6
      Top             =   135
      Width           =   540
   End
   Begin VB.Label lbl入库 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "入库时间"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblEDIT 
      AutoSize        =   -1  'True
      Caption         =   "发卡人"
      Height          =   180
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   2325
      Width           =   540
   End
   Begin VB.Label lbl至 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Index           =   3
      Left            =   2370
      TabIndex        =   3
      Top             =   105
      Width           =   180
   End
End
Attribute VB_Name = "frmBillInFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mArrFilter As Variant
Private mstrPrivs As String, mlngModule As Long
Public Event zlRefreshCon(ByVal arrFilter As Variant)
Public Event WindowsHeight(lngHeght As Long)
Private mblnNotSize  As Boolean

Private Function GetFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取条件信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-15 14:12:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
    '基本查询条件
    Set cllFilter = New Collection
    cllFilter.Add Trim(txt登记人.Text), "登记人"
    cllFilter.Add Array(Format(dtpStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "入库时间"
    cllFilter.Add IIf(chk仅有余额.Value = 1, 1, 0), "仅显示有库存发票"
    If zlDatabase.DateMoved(Format(dtpStartDate(0), "yyyy-MM-dd hh:mm:ss"), , , Me.Caption) Then
        cllFilter.Add 1, "包含历史数据"
    Else
        cllFilter.Add 0, "包含历史数据"
    End If
    Set mArrFilter = cllFilter
End Function

Private Sub cmd刷新_Click()
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter)
End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2010-11-15 14:15:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    dtpEndDate(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEndDate(0).Value = dtpEndDate(0).MaxDate
    dtpStartDate(0).Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-mm-dd")  '缺省按1月找
    
End Sub
 

Private Sub dtpEndDate_Change(index As Integer)
     If dtpEndDate(index).Value > dtpStartDate(index).MaxDate Then dtpEndDate(index).Value = dtpStartDate(index).MaxDate
    If dtpEndDate(index).Value < dtpStartDate(index).Value Then
        dtpStartDate(index).Value = dtpEndDate(index).Value
    End If
End Sub
Private Sub dtpEndDate_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_Change(index As Integer)
    If dtpStartDate(index).Value > dtpEndDate(index).MaxDate Then dtpStartDate(index).Value = dtpEndDate(index).MaxDate
    If dtpEndDate(index).Value < dtpStartDate(index).Value Then
        dtpEndDate(index).Value = dtpStartDate(index).Value
    End If
End Sub

Private Sub dtpStartDate_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub
Private Sub Form_Load()
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    chk仅有余额.Value = IIf(Val(zlDatabase.GetPara("仅显示有库存数票据", glngSys, mlngModule, 0)) = 1, 1, 0)
    If InStr(1, mstrPrivs, ";允许操作他人登记票据;") = 0 Then
        txt登记人.Text = UserInfo.姓名
        txt登记人.Enabled = False
        txt登记人.BackColor = cmd刷新.BackColor
    End If
    Call InitData
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If mblnNotSize Then Exit Sub
    mblnNotSize = True
    If ScaleWidth - (txt登记人.Width + txt登记人.Left + 50 + chk仅有余额.Width) < 0 Then
        cmd刷新.Left = txt登记人.Left
        cmd刷新.Top = dtpStartDate(0).Top + dtpStartDate(0).Height + 80
        chk仅有余额.Left = dtpStartDate(0).Left
        chk仅有余额.Top = cmd刷新.Top + (cmd刷新.Height - chk仅有余额.Height) \ 2
        RaiseEvent WindowsHeight(900)
    Else
        chk仅有余额.Left = txt登记人.Width + txt登记人.Left + 50
        chk仅有余额.Top = txt登记人.Top + (txt登记人.Height - chk仅有余额.Height) \ 2
         If ScaleWidth - (chk仅有余额.Width + chk仅有余额.Left + cmd刷新.Width + 50) < 0 Then
            cmd刷新.Left = dtpStartDate(0).Left
            cmd刷新.Top = dtpStartDate(0).Top + dtpStartDate(0).Height + 70
            RaiseEvent WindowsHeight(900)
         Else
            RaiseEvent WindowsHeight(450)
            cmd刷新.Left = chk仅有余额.Left + chk仅有余额.Width + 300
            cmd刷新.Top = dtpStartDate(0).Top - (cmd刷新.Height - dtpStartDate(0).Height) \ 2
         End If
    End If
    mblnNotSize = False
End Sub

Public Sub Init条件()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关条件
    '编制:刘兴洪
    '日期:2009-11-18 14:48:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitData
End Sub
Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
End Property

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zlDatabase.SetPara "仅显示有库存数票据", IIf(chk仅有余额.Value = 1, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub txt登记人_Change()
    txt登记人.Tag = ""
End Sub
Private Sub txt登记人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txt登记人.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txt登记人.Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select人员选择器(Me, txt登记人, Trim(txt登记人.Text)) = False Then
        Exit Sub
    End If
End Sub
Public Sub ReActionFilter()
    '重新缴活过滤
    cmd刷新_Click
End Sub



