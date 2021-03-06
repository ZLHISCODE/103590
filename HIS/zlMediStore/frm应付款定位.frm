VERSION 5.00
Begin VB.Form frm应付款定位 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应付款定位条件"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frm应付款定位.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1650
      MaxLength       =   100
      TabIndex        =   2
      Top             =   450
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -300
      TabIndex        =   9
      Top             =   1800
      Width           =   5505
   End
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   240
      Left            =   3720
      TabIndex        =   6
      Top             =   1380
      Width           =   255
   End
   Begin VB.OptionButton opt定位 
      Caption         =   "按单据号定位(&N)"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.OptionButton opt定位 
      Caption         =   "按药品供应商定位(&S)"
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2760
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2355
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "入库单据号(&M)"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   1
      Top             =   540
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "供应商(&U)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   810
      TabIndex        =   4
      Top             =   1440
      Width           =   810
   End
End
Attribute VB_Name = "frm应付款定位"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mstr单据号 As String
Dim mstr供应商ID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If StrIsValid(txtEdit(lngIndex).Text, txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                Exit Sub
            End If
            
            Select Case lngIndex
                Case 0
                    mstr供应商ID = txtEdit(lngIndex).Tag
                Case 1
                    mstr单据号 = UCase(Trim(txtEdit(lngIndex).Text))
            End Select
        End If
    Next
    
    If mstr单据号 = "" And mstr供应商ID = "" Then
        MsgBox "请输入定位条件。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd上级_Click()
    Dim rs供应商 As New ADODB.Recordset
    
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 药品供应商 Where " & _
                " nvl(撤档时间,to_date('3000-01-01','yyyy-MM-dd'))=to_date('3000-01-01','yyyy-MM-dd') " & _
                " start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    Call OpenRecordset(rs供应商, Me.Caption)
    
    If rs供应商.EOF Then
        rs供应商.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rs供应商
        .StrNode = "所有药品供应商"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = True Then
            txtEdit(0).Tag = .CurrentID
            txtEdit(0).Text = .CurrentName
        End If
    End With
    Unload FrmSelect
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function Get定位条件(str单据号 As String, str供应商ID As String) As Boolean
    
    frm应付款定位.Show vbModal, frm应付款查询
    
    Get定位条件 = mblnOK
    If mblnOK = True Then
        str单据号 = mstr单据号
        str供应商ID = mstr供应商ID
    End If
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Sub opt定位_Click(Index As Integer)
    txtEdit(0).Enabled = opt定位(0).Value
    lbl(0).Enabled = opt定位(0).Value
    cmd上级.Enabled = opt定位(0).Value
    
    txtEdit(1).Enabled = opt定位(1).Value
    lbl(1).Enabled = opt定位(1).Value
    
    txtEdit(Index).SetFocus
End Sub
