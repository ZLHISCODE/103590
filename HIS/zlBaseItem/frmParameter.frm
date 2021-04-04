VERSION 5.00
Begin VB.Form frmParameter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmParameter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4560
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPriceFolw 
      Caption         =   "调价流程"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4335
      Begin VB.CheckBox chkPriceFlow 
         Caption         =   "调价需要审核"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   0
      Top             =   1920
      Width           =   1100
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnLoad As Boolean '界面是否加载完成 true-完成 false-未完成

Public Sub ShowMe(ByVal fraParent As Form)
    Me.Show vbModal, fraParent
End Sub

Private Sub LoadData()
    Dim int调价 As Integer
    
    int调价 = zldatabase.GetPara("调价需要审核", glngSys, 1009, 0)
    chkPriceFlow.Value = IIF(int调价 = 1, 1, 0)
End Sub


Private Sub chkPriceFlow_Click()
    Dim blnResult As Boolean
    
    If mblnLoad = True Then
        blnResult = checkNotPrice
        If blnResult = False Then
            MsgBox "还存在未生效的调价单据，不能修改此参数！", vbInformation, gstrSysName
            chkPriceFlow.Value = 1
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnLoad = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    zldatabase.SetPara "调价需要审核", IIF(chkPriceFlow.Value = 1, "1", "0"), glngSys, 1009
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadData
    mblnLoad = True
End Sub

Private Function checkNotPrice() As Boolean
    '检查是否还存在未生效的价格
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If chkPriceFlow.Value = 0 Then
        gstrSQL = "Select 1 From 收费调价记录 Where 审核标志 = 0 And Rownum <= 1"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "未生效单据查询")
        If rsData.RecordCount > 0 Then
            checkNotPrice = False
        Else
            checkNotPrice = True
        End If
    Else
        checkNotPrice = True
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
