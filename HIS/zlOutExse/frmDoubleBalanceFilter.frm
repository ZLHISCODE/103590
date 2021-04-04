VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDoubleBalanceFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt住院号 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      MaxLength       =   18
      TabIndex        =   5
      Top             =   1290
      Width           =   1830
   End
   Begin VB.CheckBox chkDelRecord 
      Caption         =   "含退费结算记录"
      Height          =   345
      Left            =   3705
      TabIndex        =   26
      Top             =   255
      Width           =   2160
   End
   Begin VB.TextBox txt姓名 
      Height          =   300
      IMEMode         =   1  'ON
      Left            =   1065
      MaxLength       =   64
      TabIndex        =   2
      Top             =   915
      Width           =   1830
   End
   Begin VB.TextBox txt门诊号 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1065
      MaxLength       =   18
      TabIndex        =   4
      Top             =   1290
      Width           =   1830
   End
   Begin VB.OptionButton opt病人 
      Caption         =   "门诊病人和住院病人"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   3705
      TabIndex        =   13
      Top             =   2535
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton opt病人 
      Caption         =   "住院病人"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   2415
      TabIndex        =   12
      Top             =   2535
      Width           =   1020
   End
   Begin VB.OptionButton opt病人 
      Caption         =   "门诊病人"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   11
      Top             =   2535
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   14
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   15
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   5880
      TabIndex        =   16
      Top             =   1605
      Width           =   1100
   End
   Begin VB.ComboBox cbo操作员 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   915
      Width           =   1830
   End
   Begin VB.TextBox txtNOBegin 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1065
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1680
      Width           =   1830
   End
   Begin VB.TextBox txtNoEnd 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1680
      Width           =   1830
   End
   Begin VB.TextBox txtFactBegin 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1065
      TabIndex        =   8
      Top             =   2055
      Width           =   1830
   End
   Begin VB.TextBox txtFactEnd 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3705
      TabIndex        =   10
      Top             =   2055
      Width           =   1830
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   525
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   185925635
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1065
      TabIndex        =   0
      Top             =   105
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   185925635
      CurrentDate     =   36588
   End
   Begin VB.Label lbl住院号 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院号"
      Height          =   180
      Left            =   3090
      TabIndex        =   27
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   180
      Left            =   630
      TabIndex        =   25
      Top             =   975
      Width           =   360
   End
   Begin VB.Label lbl门诊号 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊号"
      Height          =   180
      Left            =   225
      TabIndex        =   24
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label lblFil 
      Alignment       =   1  'Right Justify
      Caption         =   "病人来源"
      Height          =   180
      Left            =   60
      TabIndex        =   23
      Top             =   2535
      Width           =   930
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      Height          =   180
      Left            =   270
      TabIndex        =   22
      Top             =   165
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间"
      Height          =   180
      Left            =   270
      TabIndex        =   21
      Top             =   585
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   3255
      TabIndex        =   20
      Top             =   1740
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单据号"
      Height          =   180
      Left            =   450
      TabIndex        =   19
      Top             =   1740
      Width           =   540
   End
   Begin VB.Label lbl操作员 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作员"
      Height          =   180
      Left            =   3090
      TabIndex        =   18
      Top             =   975
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   3255
      TabIndex        =   17
      Top             =   2115
      Width           =   180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据号"
      Height          =   180
      Left            =   450
      TabIndex        =   9
      Top             =   2115
      Width           =   540
   End
End
Attribute VB_Name = "frmDoubleBalanceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long, mstrPrivs As String, mfrmParent As frmDoubleBalanceNormal
Public mblnInit As Boolean

Public Sub InitFilter(frmMain As Object, lngModul As Long, strPrivs As String)
    Set mfrmParent = frmMain
    mlngModul = lngModul
    mstrPrivs = strPrivs
    mblnInit = True
    Me.Show vbModal, frmMain
End Sub

Public Function FilterInited() As Boolean
    FilterInited = mblnInit
End Function

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    mblnInit = True
    Call mfrmParent.ReadData(0, mstrPrivs)
    Me.Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub LoadOperator()
    Dim rsTmp As New ADODB.Recordset, i As Integer
    '操作员
    cbo操作员.Clear
    If InStr(mstrPrivs, "所有操作员") > 0 Then
        cbo操作员.AddItem "所有收费员"
        Set rsTmp = GetPersonnel("门诊收费员", True)
        For i = 1 To rsTmp.RecordCount
            cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo操作员.ItemData(cbo操作员.NewIndex) = rsTmp!ID
            If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
            rsTmp.MoveNext
        Next
    Else
        cbo操作员.AddItem UserInfo.简码 & "-" & UserInfo.姓名
        cbo操作员.ItemData(cbo操作员.NewIndex) = UserInfo.ID
    End If
    If cbo操作员.ListIndex = -1 And cbo操作员.ListCount > 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_Load()
    Dim Curdate As Date
    Call LoadOperator
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    txt姓名.Text = "": txtFactBegin.Text = "": txtFactEnd.Text = ""
    txtNOBegin.Text = "": txtNoEnd.Text = ""
    txt门诊号.Text = "": txt住院号.Text = ""
    chkDelRecord.Value = 0
    Call opt病人_Click(0)
End Sub

Private Sub opt病人_Click(Index As Integer)
    If opt病人(0).Value Then
        lbl门诊号.Visible = True
        txt门诊号.Visible = True
        lbl住院号.Visible = False
        txt住院号.Visible = False
    ElseIf opt病人(1).Value Then
        lbl门诊号.Visible = False
        txt门诊号.Visible = False
        lbl住院号.Visible = True
        txt住院号.Visible = True
        lbl住院号.Left = lbl门诊号.Left
        txt住院号.Left = txt门诊号.Left
    ElseIf opt病人(2).Value Then
        lbl门诊号.Visible = True
        txt门诊号.Visible = True
        lbl住院号.Visible = True
        txt住院号.Visible = True
        lbl住院号.Left = lbl操作员.Left
        txt住院号.Left = cbo操作员.Left
    End If
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
    
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 13)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 13)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub
