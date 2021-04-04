VERSION 5.00
Begin VB.Form frmHF 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "页眉与页脚"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmHF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   0
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "日期"
      Height          =   615
      Left            =   3180
      Picture         =   "frmHF.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmd页数 
      Caption         =   "总页数"
      Height          =   615
      Left            =   1170
      Picture         =   "frmHF.frx":06F6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdMan 
      Caption         =   "用户名"
      Height          =   615
      Left            =   4200
      Picture         =   "frmHF.frx":0DE0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "时间"
      Height          =   615
      Left            =   2190
      Picture         =   "frmHF.frx":14CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdUnit 
      Caption         =   "单位名"
      Height          =   615
      Left            =   5220
      Picture         =   "frmHF.frx":1BB4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmd页码 
      Caption         =   "页码"
      Height          =   615
      Left            =   150
      Picture         =   "frmHF.frx":229E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   1365
      Index           =   3
      Left            =   4230
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2340
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   1365
      Index           =   2
      Left            =   2250
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2340
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Height          =   1365
      Index           =   1
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2340
      Width           =   1785
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHF.frx":2988
      Height          =   705
      Left            =   270
      TabIndex        =   14
      Top             =   150
      Width           =   4125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "右(&R):"
      Height          =   180
      Left            =   4230
      TabIndex        =   12
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中(&M):"
      Height          =   180
      Left            =   2280
      TabIndex        =   10
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "左(&L):"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   2070
      Width           =   540
   End
End
Attribute VB_Name = "frmHF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'本窗体用在设置页眉和页脚




Dim mstrTemp As String      '临时的页眉页脚值
Dim mblnTemp As Boolean     '为假表示是按"取消"关闭窗口
Dim mintIndex As Integer    '获得焦点的Text1的索引值

Private Sub Form_Load()
    Dim intPos As Integer
    Dim intPos1 As Integer
    mblnTemp = False
    On Error Resume Next
    intPos = InStr(mstrTemp, ";")
    intPos1 = intPos + 1
    Text1(1).Text = Mid(mstrTemp, 1, intPos - 1)
    intPos = InStr(intPos1, mstrTemp, ";")
    Text1(2).Text = Mid(mstrTemp, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    Text1(3).Text = Mid(mstrTemp, intPos1)
    mintIndex = 1
    'On Error GoTo 0
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    mintIndex = Index
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrTemp = Text1(1).Text & ";" & Text1(2).Text & ";" & Text1(3).Text
    mblnTemp = True
    Unload Me
End Sub

Public Function GetText(strGet As String) As Boolean
    mstrTemp = strGet
    Me.Show 1
    strGet = mstrTemp
    GetText = mblnTemp
End Function

Private Sub cmd页码_Click()
    Text1(mintIndex).SelText = "第[页码]页"
End Sub

Private Sub cmd页数_Click()
    Text1(mintIndex).SelText = "共[页数]页"
End Sub

Private Sub cmdTime_Click()
    Text1(mintIndex).SelText = "[时间]"
End Sub

Private Sub cmdDate_Click()
    Text1(mintIndex).SelText = "[日期]"
End Sub

Private Sub cmdMan_Click()
    Text1(mintIndex).SelText = "[用户名]"
End Sub

Private Sub cmdUnit_Click()
    Text1(mintIndex).SelText = "[单位名]"
End Sub
