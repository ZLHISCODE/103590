VERSION 5.00
Begin VB.Form frmLabGroupEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检验小组"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "frmLabGroupEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5145
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk连续 
      Caption         =   "连续增加"
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   1275
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3405
      TabIndex        =   5
      Top             =   1215
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2145
      TabIndex        =   4
      Top             =   1215
      Width           =   1100
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   555
      Width           =   3570
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   1155
      TabIndex        =   1
      Top             =   195
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "小组名称"
      Height          =   270
      Left            =   315
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "小组编码"
      Height          =   270
      Left            =   285
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmLabGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngGroupID As Long
Private mblnAdd As Boolean
Private mblnOK As Boolean
Private mstr编码名称 As String

Public Function ShowMe(ByRef lngItemID As Long, ByVal intAdd As Integer, ByVal str编码名称 As String, ByVal frmMain As Form) As Boolean
    mlngGroupID = lngItemID
    mblnAdd = intAdd = 1
    mstr编码名称 = str编码名称
    mblnOK = False
    Me.Show vbModal, frmMain
    ShowMe = mblnOK
    lngItemID = mlngGroupID
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, str编码 As String, str名称 As String
    
    On Error GoTo ErrHandle
    str编码 = Replace(Trim(txt编码.Text), "'", "")
    str名称 = Replace(Trim(txt名称.Text), "'", "")
    If str名称 = "" Or str名称 = "" Then
        MsgBox "编码和名称不能为空！", vbInformation, Me.Caption
        Exit Sub
    End If
    If txt编码.Tag = "新增" Then
        mlngGroupID = zlDatabase.GetNextId("部门表")
        strSQL = "zl_检验小组_Edit(1," & mlngGroupID & ",'" & str编码 & "','" & str名称 & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Else
        strSQL = "zl_检验小组_Edit(2," & mlngGroupID & ",'" & str编码 & "','" & str名称 & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    mblnOK = True
    If chk连续.Value = 1 Then
        txt编码.Text = ""
        txt名称.Text = ""
    Else
        Unload Me
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If mblnAdd Then
        txt编码.Tag = "新增"
        txt编码.Text = ""
        txt名称.Text = ""
    Else
        txt编码.Tag = "编辑"
        txt编码.Text = Split(mstr编码名称, "|")(0)
        txt名称.Text = Split(mstr编码名称, "|")(1)
    End If
End Sub
