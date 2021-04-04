VERSION 5.00
Begin VB.Form frm批量设置向导 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "批量设置向导"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   Icon            =   "frm批量设置向导.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6885
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   5640
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认"
      Height          =   350
      Left            =   4200
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fra新窗口 
      Caption         =   "替换为新窗口"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   6615
      Begin VB.ComboBox cbo窗口 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   1850
      Width           =   6855
   End
   Begin VB.Frame fra类型 
      Caption         =   "选择处方类型"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6615
      Begin VB.CheckBox chk处方 
         Caption         =   "精神II类"
         Height          =   255
         Index           =   5
         Left            =   5450
         TabIndex        =   11
         Top             =   320
         Width           =   1095
      End
      Begin VB.CheckBox chk处方 
         Caption         =   "精神I类"
         Height          =   255
         Index           =   4
         Left            =   4192
         TabIndex        =   10
         Top             =   320
         Width           =   975
      End
      Begin VB.CheckBox chk处方 
         Caption         =   "麻醉"
         Height          =   255
         Index           =   3
         Left            =   3174
         TabIndex        =   9
         Top             =   320
         Width           =   735
      End
      Begin VB.CheckBox chk处方 
         Caption         =   "急诊"
         Height          =   255
         Index           =   2
         Left            =   2156
         TabIndex        =   8
         Top             =   320
         Width           =   735
      End
      Begin VB.CheckBox chk处方 
         Caption         =   "儿科"
         Height          =   255
         Index           =   1
         Left            =   1138
         TabIndex        =   7
         Top             =   320
         Width           =   735
      End
      Begin VB.CheckBox chk处方 
         Caption         =   "普通"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   320
         Width           =   735
      End
   End
   Begin VB.Frame fra现窗口 
      Caption         =   "选择现窗口"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chk窗口 
         Caption         =   "动态添加窗口"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm批量设置向导"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng药房ID As Long
Private mstr窗口 As String
Private mstrConWin As String   '选择的现窗口
Private mstr处方 As String
Private mstr新窗口 As String
Private mstr现窗口 As String

Private Sub LoadOldWin()
    Dim i As Integer
    
    If mstr窗口 = "" Then Exit Sub
    
    Me.chk窗口(0).Caption = Split(mstr窗口, ",")(0)
    chk窗口(0).Width = 150 + LenB(chk窗口(0).Caption) * 128
    For i = 1 To UBound(Split(mstr窗口, ",")) - 1
        Load chk窗口(i)
        chk窗口(i).Visible = True
        chk窗口(i).Caption = Split(mstr窗口, ",")(i)
        chk窗口(i).Width = 150 + LenB(chk窗口(i - 1).Caption) * 128
        chk窗口(i).Left = chk窗口(i - 1).Left + chk窗口(i - 1).Width + 100
    Next
End Sub
Private Sub LoadNewWin()
    Dim strSql As String
    Dim rsRecord As Recordset
    
    On Error GoTo errHandle
    
    strSql = "select 编码,名称 from 发药窗口 where 药房id=[1] and 上班否=1"
    If mstr现窗口 <> "" Then strSql = strSql & " And 名称<>[2] "
    Set rsRecord = zldatabase.OpenSQLRecord(strSql, "Init窗口", mlng药房ID, mstr现窗口)
    
    If Not (rsRecord Is Nothing) Then
        Do While Not rsRecord.EOF
            Me.cbo窗口.AddItem rsRecord!名称
            mstr窗口 = mstr窗口 & rsRecord!名称 & ","
            rsRecord.MoveNext
        Loop
    End If
    
    If mstr现窗口 <> "" Then mstr窗口 = mstr现窗口
    
    Me.cbo窗口.ListIndex = 0
    Exit Sub
errHandle:
    If errcenter() = 1 Then Resume
    Call saveerrlog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim i As Integer
    For i = 0 To Me.chk窗口.UBound
        If Me.chk窗口(i).Value = 1 Then
            mstrConWin = mstrConWin & Me.chk窗口(i).Caption & ","
        End If
    Next
    
    For i = 0 To Me.chk处方.UBound
        If Me.chk处方(i).Value = 1 Then
            mstr处方 = mstr处方 & Me.chk处方(i).Caption & ","
        End If
    Next
    
    mstr新窗口 = Me.cbo窗口.Text
    
    Unload Me
End Sub

Private Sub Form_Load()
    LoadNewWin
    LoadOldWin
End Sub

Public Sub ShowME(ByVal lng药房ID As Long, ByRef strConWin As String, ByRef str处方 As String, ByRef str新窗口 As String, Optional str现窗口 As String = "")
    mlng药房ID = lng药房ID
    mstr现窗口 = str现窗口
    Me.Show 1
    
    strConWin = mstrConWin
    str处方 = mstr处方
    str新窗口 = mstr新窗口
    
    
    mstrConWin = ""
    mstr处方 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr窗口 = ""
End Sub
