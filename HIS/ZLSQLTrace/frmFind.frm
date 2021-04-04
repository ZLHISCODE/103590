VERSION 5.00
Begin VB.Form frmFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin ZLSQLTrace.ccXPButton cmdClose 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3630
      TabIndex        =   4
      Top             =   930
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   "取消(&C)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ZLSQLTrace.ccXPButton cmdFind 
      Default         =   -1  'True
      Height          =   345
      Left            =   2625
      TabIndex        =   3
      Top             =   930
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   "确定(&O)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "区分大小写(&M)"
      Height          =   195
      Left            =   1005
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   1470
   End
   Begin VB.ComboBox cboFind 
      Height          =   300
      Left            =   1005
      TabIndex        =   1
      Top             =   225
      Width           =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找内容"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   285
      Width           =   720
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Find(ByVal Text As String, ByVal MatchCase As Boolean)

Public Sub ShowMe(ByVal Text As String)
    If Text <> "" Then
        Text = Split(Text, vbCrLf)(0)
        If Len(Text) > 100 Then Text = Left(Text, 100)
    End If
    
    cboFind.Text = Text
    Me.Show , frmMain
End Sub

Private Sub cboFind_GotFocus()
    cboFind.SelStart = 0: cboFind.SelLength = (Len(cboFind.Text))
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 32 Or KeyAscii < 0 Then
        Call CboAppendText(cboFind, KeyAscii)
    End If
End Sub

Private Sub cmdClose_Click()
    Hide
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String, i As Integer
    
    If cboFind.Text = "" Then Exit Sub
    strFind = cboFind.Text
    
    For i = 0 To cboFind.ListCount - 1
        If cboFind.List(i) = strFind Then
            cboFind.RemoveItem i: Exit For
        End If
    Next
    cboFind.AddItem strFind, 0
    cboFind.ListIndex = cboFind.NewIndex
    
    Call cboFind_GotFocus: cboFind.SetFocus
    RaiseEvent Find(cboFind.Text, chkCase.Value = 1)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cboFind.SetFocus
    cboFind_GotFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer, k As Integer, s As String
        
    s = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "FindFormLocate", "")
    If s = "" Then
        Me.Left = frmMain.Left + (frmMain.Width - Me.Width) * (2 / 3)
        Me.Top = frmMain.Top + (frmMain.Height - Me.Height) * (1 / 3)
    Else
        Me.Left = frmMain.Left + Split(s, ",")(0)
        Me.Top = frmMain.Top + Split(s, ",")(1)
    End If
        
    k = Val(GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "FindCount", 0))
    For i = 1 To k
        s = GetSetting("ZLSOFT\公共模块\ZLDBATools", "FindItem", "Find" & i, "")
        If s <> "" Then cboFind.AddItem s
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, x As Variant
        
    x = GetAllSettings("ZLSOFT\公共模块\ZLDBATools", "FindItem")
    If IsArray(x) Then DeleteSetting "ZLSOFT\公共模块\ZLDBATools", "FindItem"

    SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "FindCount", cboFind.ListCount
    For i = 0 To cboFind.ListCount - 1
        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "FindItem", "Find" & i + 1, cboFind.List(i)
    Next
    
    SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "FindFormLocate", Me.Left - frmMain.Left & "," & Me.Top - frmMain.Top
End Sub
