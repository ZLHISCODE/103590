VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   13395
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "关闭"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   13095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'显示界面函数
'   hwnd:父窗口句柄
'   nCmdShow:显示窗口命令
'返回:true(成功);false(失败)
Private Declare Function CEC_ShowWindows Lib "E:\CecDeviceToHis.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean

'退出动态库函数
'返回:true(成功);false(失败)
Private Declare Function CEC_Uninitialize Lib "E:\CecDeviceToHis.dll" () As Boolean

Private Sub Command1_Click()
   CEC_ShowWindows 0, 3
End Sub

Private Sub Command2_Click()
   CEC_ShowWindows 0, 0
End Sub

Private Sub Form_Load()
    Initialize Frame1.hwnd
End Sub

Private Sub Form_Close()
    CEC_Uninitialize
End Sub

