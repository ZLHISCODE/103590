VERSION 5.00
Begin VB.Form frm等待返回 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm等待返回.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdGet 
      Caption         =   "读取计算结果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1500
      TabIndex        =   2
      Top             =   1620
      Width           =   2595
   End
   Begin VB.Timer timRead 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   510
      Top             =   1650
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   127
      TabIndex        =   0
      Top             =   90
      Width           =   5490
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "正在等待医保接口返回计算结果……"
         BeginProperty Font 
            Name            =   "仿宋_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   330
         TabIndex        =   1
         Top             =   600
         Width           =   4800
      End
   End
End
Attribute VB_Name = "frm等待返回"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsReturn As ADODB.Recordset
Dim mstrSQL As String
Private mblnOK As Boolean

Public Function WaitForYB(rsReturn As ADODB.Recordset, ByVal strSQL As String) As Boolean
    mblnOK = False
    
    Set mrsReturn = rsReturn
    mstrSQL = strSQL
    
    timRead.Enabled = False
    frm等待返回.Show vbModal
    WaitForYB = mblnOK
    Set rsReturn = mrsReturn
End Function

Private Sub cmdGet_Click()
    timRead.Enabled = True
End Sub

Private Sub timRead_Timer()
    If mrsReturn.State = adStateOpen Then mrsReturn.Close
    mrsReturn.Open mstrSQL, gcn咸阳, adOpenStatic, adLockReadOnly
    If mrsReturn.EOF = False Then
        '取得前置服务器返回数值，继续执行
        timRead.Enabled = False
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If MsgBox("你不再等待医保结算的结果吗？", vbYesNo + vbQuestion + vbDefaultButton2, "医保取消提示") = vbYes Then
            timRead.Enabled = False
            Unload Me
        End If
    End If
        
End Sub
