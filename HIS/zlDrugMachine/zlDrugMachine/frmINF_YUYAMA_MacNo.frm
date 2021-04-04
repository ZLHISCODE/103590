VERSION 5.00
Begin VB.Form frmINF_YUYAMA_MacNo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "选择机器"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd 
      Caption         =   "机器&2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "机器&1"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注意：关闭窗体默认选择“机器1”！"
      Height          =   180
      Left            =   795
      TabIndex        =   2
      Top             =   2280
      Width           =   2970
   End
End
Attribute VB_Name = "frmINF_YUYAMA_MacNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintMacNo As Integer

Public Function ShowMe(ByRef intMacNo As Integer) As Boolean
    Me.Show vbModal
    intMacNo = mintMacNo
End Function

Private Sub cmd_Click(Index As Integer)
    mintMacNo = Index + 1
    Unload Me
End Sub

Private Sub Form_Load()
    mintMacNo = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdlDefine.gtypYUYAMA.MacNO = mintMacNo
    If mintMacNo <= 0 Then mintMacNo = 1    '默认机器1
End Sub
