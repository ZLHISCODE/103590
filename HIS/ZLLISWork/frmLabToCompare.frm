VERSION 5.00
Begin VB.Form frmLabToCompare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "置为比对标本"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5730
   Icon            =   "frmLabToCompare.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frmLine 
      Height          =   30
      Left            =   -30
      TabIndex        =   7
      Top             =   1650
      Width           =   5745
   End
   Begin VB.ComboBox cbo比对号 
      Height          =   300
      Left            =   2265
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3195
      TabIndex        =   1
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   0
      Top             =   1770
      Width           =   1100
   End
   Begin VB.Label lbl比对号 
      AutoSize        =   -1  'True
      Caption         =   "将样本设置为"
      Height          =   180
      Left            =   1140
      TabIndex        =   6
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label lbl检验仪器 
      AutoSize        =   -1  'True
      Caption         =   "检验仪器：####"
      Height          =   180
      Left            =   1140
      TabIndex        =   4
      Top             =   795
      Width           =   1260
   End
   Begin VB.Label lbl检验日期 
      AutoSize        =   -1  'True
      Caption         =   "检验日期：####"
      Height          =   180
      Left            =   1140
      TabIndex        =   3
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "如果认为该样本具备比对特征，请选择要指定比对号后确定。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   135
      Width           =   4860
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   210
      Picture         =   "frmLabToCompare.frx":058A
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmLabToCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long   '当前id
Private mblnOK As Boolean

'临时变量
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Public Function ShowMe(ByVal frmParent As Form, lngID As Long) As Boolean
    Dim strDate As String
    mlngID = lngID
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select L.医嘱id, To_Char(L.检验时间, 'yyyy-mm-dd') As 检验时间, M.名称 As 仪器" & vbNewLine & _
            "From 检验标本记录 L, 检验仪器 M" & vbNewLine & _
            "Where L.仪器id = M.ID And L.ID = [1] And L.检验时间 Is Not Null"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    With rsTemp
        If .RecordCount <= 0 Then
            MsgBox "该项目为手工项目或该样本尚未填写结果，不能设置为比对样本！", vbInformation, gstrSysName
            Unload Me: ShowMe = False: Exit Function
        End If
        Me.lbl检验日期.Caption = "检验日期：" & Format(!检验时间, "yyyy-MM-dd")
        Me.lbl检验仪器.Caption = "检验仪器：" & !仪器
    End With
    
    With Me.cbo比对号
        .Clear: .AddItem "比对1": .AddItem "比对2": .AddItem "比对3":: .AddItem "比对4": .AddItem "比对5": .ListIndex = 0
    End With
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False: Exit Function
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdOK_Click()

    gstrSql = "Zl_检验标本记录_置为比对(" & mlngID & "," & Me.cbo比对号.ListIndex + 1 & ")"
    Err = 0: On Error GoTo ErrHand
    zldatabase.ExecuteProcedure gstrSql, Me.Caption
    mblnOK = True: Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Call zlCommFun.OpenIme(False)
End Sub


