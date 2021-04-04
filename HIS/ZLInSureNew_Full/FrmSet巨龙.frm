VERSION 5.00
Begin VB.Form FrmSet巨龙 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "FrmSet巨龙.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt服务网点编号 
      Height          =   300
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   3045
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   5
      Top             =   1080
      Width           =   1100
   End
   Begin VB.ComboBox Cbo操作模式 
      Height          =   300
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3045
   End
   Begin VB.Label Lbl服务网点编号 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "网点编号"
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Lbl操作模式 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作模式"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmSet巨龙"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng险类 As Long
Private blnOK As Boolean

Public Function ShowSet(ByVal lng险类 As Long) As Boolean
    blnOK = False
    mlng险类 = lng险类
    
    Me.Show 1
    ShowSet = blnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_保险参数_Delete(" & mlng险类 & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & mlng险类 & ",NULL,'操作模式'," & Cbo操作模式.ListIndex & ",1)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & mlng险类 & ",NULL,'服务网点编号','" & Txt服务网点编号.Text & "',2)", , adCmdStoredProc
    gcnOracle.CommitTrans
    
    blnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim intValue As Integer
    
    '装入初始数据
    Cbo操作模式.Clear
    Cbo操作模式.AddItem "先办理出院结算,再办理出院手续"
    Cbo操作模式.AddItem "先办理出院手续,再办理出院结算"
    
    '获取参数值
    intValue = 0
    gstrSQL = "Select Nvl(参数值,0) Value From 保险参数 Where 险类=[1] And 参数名='操作模式'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取参数值", mlng险类)
    
    If Not rsTmp.EOF Then
        intValue = rsTmp!Value
    End If
    Cbo操作模式.ListIndex = intValue
    
    '服务网点编号
    gstrSQL = "Select Nvl(参数值,'') Value From 保险参数 Where 险类=[1] And 参数名='服务网点编号'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取参数值", mlng险类)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!Value) Then
            Txt服务网点编号.Text = rsTmp!Value
        End If
    End If
End Sub


