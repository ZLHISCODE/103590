VERSION 5.00
Begin VB.Form frmPriceGo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定位设置"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   4770
      Begin VB.OptionButton optHead 
         Caption         =   "所有"
         Height          =   195
         Left            =   2985
         TabIndex        =   3
         Top             =   735
         Width           =   660
      End
      Begin VB.OptionButton optCur 
         Caption         =   "向下"
         Height          =   195
         Left            =   3705
         TabIndex        =   4
         Top             =   735
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3105
         MaxLength       =   100
         TabIndex        =   1
         Top             =   255
         Width           =   1275
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   1365
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         MaxLength       =   8
         TabIndex        =   0
         Top             =   255
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   2700
         TabIndex        =   10
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "划价人"
         Height          =   180
         Left            =   345
         TabIndex        =   9
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   345
         TabIndex        =   8
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2400
      TabIndex        =   5
      Top             =   1305
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   1305
      Width           =   1100
   End
End
Attribute VB_Name = "frmPriceGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub



Private Sub cmdOK_Click()
    If txtNO.Text = "" And cbo操作员.ListIndex = 0 And txt姓名.Text = "" Then
        MsgBox "请至少设定一个条件！", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
   '问题:30532
    If InStr(1, txtNO.Text, "[") > 0 Then
        MsgBox "单据号中含用非法字符[]", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
    If InStr(1, txtNO.Text, "]") > 0 Then
        MsgBox "单据号中含用非法字符[]", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
    If InStr(1, txt姓名.Text, "[") > 0 Then
        MsgBox "姓名中含用非法字符[]", vbInformation, gstrSysName
        txt姓名.SetFocus: Exit Sub
    End If
    If InStr(1, txt姓名.Text, "]") > 0 Then
        MsgBox "姓名中含用非法字符[]", vbInformation, gstrSysName
        txt姓名.SetFocus: Exit Sub
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    txtNO.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If InStr(1, "[]", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim bln仅操作员部门 As Boolean
    
    gblnOK = False

    cbo操作员.AddItem ""
    cbo操作员.ListIndex = 0
        
    bln仅操作员部门 = zlstr.IsHavePrivs(mstrPrivs, "所有科室") = False And gblnUserIsClinic '113577
    Set rsTmp = GetPersonnel("药房发药人,医生,护士", True, bln仅操作员部门) '113646
    For i = 1 To rsTmp.RecordCount
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        rsTmp.MoveNext
    Next
    cbo.SetListWidth cbo操作员.hWnd, cbo操作员.Width * 3 / 2
End Sub

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNO, KeyAscii, m文本式
End Sub

Private Sub txtNO_LostFocus()
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 13)
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub
