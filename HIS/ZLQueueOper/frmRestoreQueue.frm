VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRestoreQueue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "重排"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   Icon            =   "frmRestoreQueue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComCtl2.DTPicker dtpQueueDate 
      Height          =   315
      Left            =   1185
      TabIndex        =   5
      Top             =   660
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Format          =   241827841
      CurrentDate     =   41650
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&S)"
      Height          =   400
      Left            =   2010
      TabIndex        =   3
      Top             =   1140
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   3195
      TabIndex        =   2
      Top             =   1140
      Width           =   1100
   End
   Begin VB.ComboBox cbxQueueName 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "排队日期："
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   705
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "队列名称："
      Height          =   210
      Left            =   195
      TabIndex        =   0
      Top             =   285
      Width           =   960
   End
End
Attribute VB_Name = "frmRestoreQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean


Public Sub ShowRestoreQueueWindow(ByVal strQueueNames As String, _
                                        ByVal strCurQueueName As String, _
                                        ByRef strNewQueueName As String, _
                                        ByVal dtCurQueueDate As Date, _
                                        ByRef dtQueueDate As Date, _
                                        Optional objOwner As Object)
'显示重排队列窗口
    mblnOk = False
    
    strNewQueueName = ""
    dtQueueDate = dtCurQueueDate
    
    dtpQueueDate.value = dtCurQueueDate
    
    Call LoadQueueName(strQueueNames, strCurQueueName)
    
    Me.Show 1, objOwner
    
    If mblnOk = True Then
        strNewQueueName = Me.cbxQueueName.Text
        dtQueueDate = Me.dtpQueueDate.value
    End If
End Sub


Private Sub LoadQueueName(ByVal strQueueNames As String, ByVal strCurQueueName As String)
'载入队列名称
    Dim strQueueName() As String
    Dim i As Long
    
    strQueueName = Split(IIf(Trim(strQueueNames) = "", strCurQueueName, strQueueNames) & ",", ",")
    
    cbxQueueName.Clear
    
    For i = 0 To UBound(strQueueName)
        If Trim(strQueueName(i)) <> "" Then
            cbxQueueName.AddItem strQueueName(i)
            
            If Trim(strQueueName(i)) = strCurQueueName Then
                cbxQueueName.ListIndex = i
            End If
        End If
    Next i
    
    If cbxQueueName.ListCount > 0 And cbxQueueName.ListIndex < 0 Then cbxQueueName.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If dtpQueueDate.value < CDate(Format(Now, "yyyy-mm-dd")) Then
        MsgBox "排队日期不能小于当天日期。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    mblnOk = True
    
    Me.Hide
End Sub

