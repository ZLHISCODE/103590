VERSION 5.00
Begin VB.Form frmRadMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "影像项目修改"
   ClientHeight    =   4080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6150
   Icon            =   "frmRadMod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6150
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   15
      Top             =   3615
      Width           =   1100
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   195
      Picture         =   "frmRadMod.frx":058A
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3615
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3750
      TabIndex        =   13
      Top             =   3615
      Width           =   1100
   End
   Begin VB.TextBox txt图象 
      Height          =   300
      Left            =   4425
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2235
      Width           =   780
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   11
      Top             =   3495
      Width           =   6210
   End
   Begin VB.TextBox txt准备 
      Height          =   300
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   10
      Top             =   3045
      Width           =   4230
   End
   Begin VB.ComboBox cbo胶片 
      Height          =   300
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2670
      Width           =   2055
   End
   Begin VB.ComboBox cbo病检 
      Height          =   300
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2295
      Width           =   2055
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1935
      Width           =   2055
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   1
      Top             =   510
      Width           =   6210
   End
   Begin VB.Label lblPartUnit 
      AutoSize        =   -1  'True
      Caption         =   "部位:     计算单位:"
      Height          =   180
      Left            =   810
      TabIndex        =   19
      Top             =   1245
      Width           =   1710
   End
   Begin VB.Label lblCodeName 
      AutoSize        =   -1  'True
      Caption         =   "编码:     名称:"
      Height          =   180
      Left            =   810
      TabIndex        =   18
      Top             =   945
      Width           =   1350
   End
   Begin VB.Label lblBaseInfo 
      AutoSize        =   -1  'True
      Caption         =   "项目名称信息："
      Height          =   180
      Left            =   630
      TabIndex        =   17
      Top             =   675
      Width           =   1260
   End
   Begin VB.Label lblExtendInfo 
      AutoSize        =   -1  'True
      Caption         =   "影像检查补充信息："
      Height          =   180
      Left            =   630
      TabIndex        =   16
      Top             =   1650
      Width           =   1620
   End
   Begin VB.Label lbl图象 
      AutoSize        =   -1  'True
      Caption         =   "报告最大图象数目"
      Height          =   180
      Left            =   4410
      TabIndex        =   7
      Top             =   1995
      Width           =   1440
   End
   Begin VB.Label lbl准备 
      AutoSize        =   -1  'True
      Caption         =   "检查准备"
      Height          =   180
      Left            =   810
      TabIndex        =   6
      Top             =   3105
      Width           =   720
   End
   Begin VB.Label lbl胶片 
      AutoSize        =   -1  'True
      Caption         =   "可发胶片"
      Height          =   180
      Left            =   810
      TabIndex        =   5
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label lbl病检 
      AutoSize        =   -1  'True
      Caption         =   "可行病检"
      Height          =   180
      Left            =   810
      TabIndex        =   4
      Top             =   2370
      Width           =   720
   End
   Begin VB.Label lbl类别 
      AutoSize        =   -1  'True
      Caption         =   "影像类别"
      Height          =   180
      Left            =   810
      TabIndex        =   2
      Top             =   1995
      Width           =   720
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   60
      Picture         =   "frmRadMod.frx":06D4
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    只能修改影像检查项目的补充信息，如修改项目名称相关信息请在诊疗项目管理中进行。"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   5325
   End
End
Attribute VB_Name = "frmRadMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem

Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Private Sub cbo病检_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo胶片_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strDescribe As String
    
    strDescribe = "'" & Split(Me.cbo类别.Text, "-")(0) & "'"
    strDescribe = strDescribe & "," & Left(Me.cbo病检.Text, 1)
    strDescribe = strDescribe & "," & Left(Me.cbo胶片.Text, 1)
    strDescribe = strDescribe & ",'" & Trim(Me.txt准备.Text) & "'"
    strDescribe = strDescribe & "," & Val(Me.txt图象.Text)
    
    gstrSql = "zl_影像检查项目_Update(" & Me.lblBaseInfo.Tag & "," & strDescribe & ")"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub cmd帮助_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Activate()
    gstrSql = "Select I.ID,I.编码, I.名称,I.标本部位, I.计算单位,R.可行病检,R.可发胶片,R.报告图象,R.检查准备" & _
            "  From 诊疗项目目录 I, 影像检查项目 R" & _
            " Where I.ID = R.诊疗项目id And I.ID=[1] "
    Err = 0: On Error GoTo ErrHand
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblBaseInfo.Tag))
        
    With rsTemp
        If .RecordCount = 0 Then MsgBox "在你修改的同时，该项目已经被删除！", vbInformation, gstrSysName: Unload Me: Exit Sub
        Me.lblCodeName.Caption = "编码:" & !编码 & "    名称:" & !名称
        Me.lblPartUnit.Caption = "部位:" & IIf(IsNull(!标本部位), "", !标本部位) & "    计算单位:" & IIf(IsNull(!计算单位), "", !计算单位)
        Me.cbo病检.ListIndex = IIf(IsNull(!可行病检), 0, !可行病检)
        Me.cbo胶片.ListIndex = IIf(IsNull(!可发胶片), 0, !可发胶片)
        Me.txt图象.Text = IIf(IsNull(!报告图象), 0, !报告图象)
        Me.txt准备.Text = IIf(IsNull(!检查准备), "", !检查准备)
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    gstrSql = "Select * From 影像检查类别 Order By 排列"
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.cbo类别.Clear
    With rsTemp
        Do While Not .EOF
            Me.cbo类别.AddItem !编码 & "-" & !名称
            If !编码 = Mid(frmRadLists.lvwKind.SelectedItem.Key, 2) Then
                Me.cbo类别.ListIndex = Me.cbo类别.NewIndex
            End If
            .MoveNext
        Loop
    End With
        
    aryTemp = Split("0-不可能;1-必须;2-选择进行", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo病检.AddItem aryTemp(intCount)
    Next
    Me.cbo病检.ListIndex = 0
    
    aryTemp = Split("0-不可能;1-必须;2-选择发放", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo胶片.AddItem aryTemp(intCount)
    Next
    Me.cbo胶片.ListIndex = 0
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt图象_GotFocus()
    Me.txt图象.SelStart = 0: Me.txt图象.SelLength = 100
End Sub

Private Sub txt图象_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt准备_GotFocus()
    Me.txt准备.SelStart = 0: Me.txt准备.SelLength = Me.txt准备.MaxLength
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt准备_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt准备_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub
