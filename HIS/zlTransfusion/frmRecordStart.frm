VERSION 5.00
Begin VB.Form frmRecordStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择开始时间"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecordStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2415
      TabIndex        =   3
      Top             =   1245
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   615
      TabIndex        =   2
      Top             =   1245
      Width           =   1100
   End
   Begin VB.TextBox txtStart 
      Height          =   300
      Left            =   1005
      TabIndex        =   1
      Top             =   255
      Width           =   2670
   End
   Begin VB.ComboBox cboOper 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   615
      Width           =   2670
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作人员"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   660
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frmRecordStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSelect As String
Private mstrOper As String
Private mstrDate As String
Private mblnOk As Boolean

Public Function ShowSelect(ByVal strSeleOper As String, ByRef strDate As String, ByRef strOper As String) As Boolean
    mstrSelect = strSeleOper
    mstrDate = strDate
    mstrOper = strOper
    mblnOk = False
    
    Me.Show vbModal
    ShowSelect = mblnOk
    If mblnOk Then
        strDate = mstrDate
        strOper = mstrOper
    End If
End Function



Private Sub cmdCancle_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim strDate As String, strOper As String
    If Not IsDate(txtStart) Then
        MsgBox "开始时间格式不对！"
        Exit Sub
    End If
    
    mstrOper = cboOper.List(cboOper.ListIndex)
    mstrDate = Format(CDate(txtStart), "yyyy-MM-dd HH:mm:ss")
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim varTmp As Variant, intIndex As Integer, i As Integer
    Me.cboOper.Clear
    varTmp = Split(mstrSelect, "|")
    
    intIndex = 0
    For i = LBound(varTmp) To UBound(varTmp)
        Me.cboOper.AddItem varTmp(i)
        If mstrOper = varTmp(i) Then
            intIndex = cboOper.NewIndex
        End If
    Next
    If cboOper.ListCount > 0 Then cboOper.ListIndex = intIndex
    
    Me.txtStart = mstrDate
    
End Sub
