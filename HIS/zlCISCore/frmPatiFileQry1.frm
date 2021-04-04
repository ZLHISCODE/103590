VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiFileQry1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历查询条件"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmPatiFileQry1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraBase 
      Height          =   2790
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   5460
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   300
         Index           =   0
         Left            =   1215
         TabIndex        =   12
         Top             =   1695
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   83755011
         CurrentDate     =   38004
      End
      Begin VB.TextBox txtAge 
         Height          =   300
         Index           =   1
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   8
         Top             =   945
         Width           =   600
      End
      Begin VB.TextBox txtAge 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   7
         Top             =   945
         Width           =   600
      End
      Begin VB.ComboBox cmbSex 
         Height          =   300
         ItemData        =   "frmPatiFileQry1.frx":000C
         Left            =   4095
         List            =   "frmPatiFileQry1.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   582
         Width           =   1080
      End
      Begin VB.ComboBox cmbFileType 
         Height          =   300
         ItemData        =   "frmPatiFileQry1.frx":0028
         Left            =   1215
         List            =   "frmPatiFileQry1.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1800
      End
      Begin VB.TextBox txtContent 
         Height          =   600
         Left            =   1215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2085
         Width           =   3930
      End
      Begin VB.TextBox txtName 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   3
         Top             =   585
         Width           =   1785
      End
      Begin VB.TextBox txtDoctor 
         Height          =   300
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1320
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   300
         Index           =   1
         Left            =   3360
         TabIndex        =   13
         Top             =   1695
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   83755011
         CurrentDate     =   38004
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   3100
         TabIndex        =   22
         Top             =   1755
         Width           =   180
      End
      Begin VB.Label lblDate 
         Caption         =   "书写日期(&R)"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   2040
         TabIndex        =   21
         Top             =   1005
         Width           =   180
      End
      Begin VB.Label lblAge 
         Caption         =   "年龄(&A)"
         Height          =   255
         Left            =   525
         TabIndex        =   6
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label lblSex 
         Caption         =   "性别(&S)"
         Height          =   255
         Left            =   3405
         TabIndex        =   4
         Top             =   645
         Width           =   735
      End
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         Caption         =   "病历内容(&M)"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   2130
         Width           =   990
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病历种类(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人姓名(&N)"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "书写医生(&D)"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   1380
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3315
      TabIndex        =   16
      Top             =   3855
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   17
      Top             =   3855
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmPatiFileQry1.frx":0077
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   840
      TabIndex        =   20
      Top             =   240
      Width           =   4620
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPatiFileQry1.frx":013D
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmPatiFileQry1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTemp As String

Public Sub GetQueryString(ByVal frmParent As Object, strQuery As String)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '查询条件串：病历种类||病人姓名||性别||最小年龄||最大年龄||医生||日期下限||日期上限||病历内容
    '---------------------------------------------------
    Dim aryTemp() As String
    
    On Error Resume Next
    If Len(Trim(strQuery)) > 0 Then
        aryTemp = Split(strQuery, "||")
        
        If Len(aryTemp(0)) = 0 Then
            Me.cmbFileType.ListIndex = 0
        Else
            Me.cmbFileType.ListIndex = CInt(aryTemp(0))
        End If
        Me.txtName = aryTemp(1)
        Me.cmbSex.Text = IIf(Len(aryTemp(2)) = 0, " ", aryTemp(2))
        Me.txtAge(0) = aryTemp(3)
        Me.txtAge(1) = aryTemp(4)
        Me.txtDoctor = aryTemp(5)
        If Len(aryTemp(6)) = 0 Then
            Me.dtDate(0).Value = DateAdd("yyyy", -1, Date)
        Else
            Me.dtDate(0).Value = CDate(aryTemp(6))
        End If
        If Len(aryTemp(7)) = 0 Then
            Me.dtDate(1).Value = Date
        Else
            Me.dtDate(1).Value = CDate(aryTemp(7))
        End If
        Me.txtContent = aryTemp(8)
    End If
    '显示窗体
    
    strTemp = ""
    Me.Show vbModal, frmParent
    strQuery = strTemp
End Sub

Private Sub cmbFileType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmbSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    '一般特性检查
    On Error Resume Next
    If Len(Trim(Me.txtAge(0))) > 0 And Len(Trim(Me.txtAge(1))) > 0 Then
        If CInt(Me.txtAge(0)) > CInt(Me.txtAge(1)) Then MsgBox "年龄范围错！", vbExclamation, gstrSysName: Me.txtAge(0).SetFocus: Exit Sub
    End If
    If Me.dtDate(0).Value > Me.dtDate(1).Value Then MsgBox "日期范围错！", vbExclamation, gstrSysName: Me.dtDate(0).SetFocus: Exit Sub
    
    strTemp = CStr(Me.cmbFileType.ListIndex) + "||" + Trim(Me.txtName) + "||" + Trim(Me.cmbSex.Text) + "||" + _
        Trim(Me.txtAge(0)) + "||" + Trim(Me.txtAge(1)) + "||" + Trim(Me.txtDoctor) + "||" + _
        Format(Me.dtDate(0).Value, "YYYY-MM-DD") + "||" + Format(Me.dtDate(1).Value, "YYYY-MM-DD") + "||" + _
        Trim(Me.txtContent)
    
    Unload Me
End Sub

Private Sub dtDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtAge_GotFocus(Index As Integer)
    With Me.txtAge(Index)
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAge_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or ifEditKey(KeyAscii)) Or Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtContent_GotFocus()
    With Me.txtContent
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtContent_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtDoctor_GotFocus()
    With Me.txtDoctor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtDoctor_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtName_GotFocus()
    With Me.txtName
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function
