VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   1485
      TabIndex        =   5
      Top             =   750
      Width           =   2835
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   1380
      TabIndex        =   4
      Top             =   1380
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   2925
      TabIndex        =   3
      Top             =   1380
      Width           =   1380
   End
   Begin VB.ComboBox cbxFilter 
      Height          =   300
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   315
      Width           =   2835
   End
   Begin VB.Label Label2 
      Caption         =   "过 滤 值："
      Height          =   270
      Left            =   555
      TabIndex        =   2
      Top             =   780
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "过滤条件："
      Height          =   270
      Left            =   570
      TabIndex        =   0
      Top             =   330
      Width           =   945
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mstrWhereFields As String



Public Sub ShowFilterWindow(ByVal strWhereFields As String, ByRef strFilterWhere As String, ByRef strValue As String, objOwner As Object)
    mblnOk = False
    strFilterWhere = ""
    strValue = ""
    
    mstrWhereFields = strWhereFields
    
    Me.Show 1, objOwner
    
    If mblnOk Then
        strFilterWhere = Me.cbxFilter.Text
        strValue = Me.txtValue.Text
    End If
End Sub


Private Sub cmdCancel_Click()
    mblnOk = False
    
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    mblnOk = True
    
    Me.Hide
End Sub



Private Sub LoadFilterWhere()
'载入过滤条件
    Dim strWhere() As String
    Dim i As Long
    Dim lngIndex As Long
    
    strWhere = Split(mstrWhereFields & ",", ",")
    
    lngIndex = -1
    For i = 0 To UBound(strWhere)
        If Trim(strWhere(i)) <> "" Then
            Me.cbxFilter.AddItem strWhere(i)
            
            If strWhere(i) = "排队号码" Then
                lngIndex = i
            End If
            
            If lngIndex < 0 And strWhere(i) = "排队号码" Then
                lngIndex = i
            End If
        End If
    Next i
    
    If lngIndex = -1 Then lngIndex = 0
    
    cbxFilter.ListIndex = lngIndex
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call LoadFilterWhere
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
