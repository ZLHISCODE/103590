VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTimeSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "条件设置"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmTimeSel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   -30
      TabIndex        =   8
      Top             =   1995
      Width           =   5115
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   7
      Top             =   690
      Width           =   4995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   5
      Top             =   2205
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2565
      TabIndex        =   4
      Top             =   2205
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   1395
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   73072643
      CurrentDate     =   36279
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   975
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   73072643
      CurrentDate     =   36279
      MinDate         =   2
   End
   Begin VB.Label lblTimeStop 
      AutoSize        =   -1  'True
      Caption         =   "结束时间(&E)"
      Height          =   180
      Left            =   840
      TabIndex        =   2
      Top             =   1455
      Width           =   990
   End
   Begin VB.Label lblTimeStart 
      AutoSize        =   -1  'True
      Caption         =   "开始时间(&B)"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   1035
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmTimeSel.frx":000C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "设置相关的过滤条件."
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   345
      Width           =   1710
   End
End
Attribute VB_Name = "frmTimeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mstrBegin As String, mstrEnd As String
Dim mstrData As String

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "开始时间大于结束时间了。", vbExclamation, gstrSysName
        Exit Sub
    End If
    mstrBegin = Format(dtpBegin.Value, "yyyy-mm-dd")
    mstrEnd = Format(dtpEnd.Value, "yyyy-mm-dd")
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function GetTimeScope(strStartDate As String, strEndDate As String, ByVal frmOwner As Form) As Boolean
'--------------------------------------------------------------
'功能：获取应付款查询日期范围及其他条件
'参数：datBegin---------起始日期
'      datEnd-----------结束日期
'      frmOwner---------调用窗体
'返回：是否查询
'说明：
'--------------------------------------------------------------
    Dim intTemp As Long
    mstrBegin = strStartDate
    mstrEnd = strEndDate
    
    If IsDate(mstrBegin) Then
        dtpBegin.Value = CDate(mstrBegin)
    Else
        dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End If
    If IsDate(mstrEnd) Then
        dtpEnd.Value = CDate(mstrEnd)
    Else
        dtpEnd.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End If
    
    
    frmTimeSel.Show vbModal, frmOwner
    GetTimeScope = mblnOK
    If mblnOK = True Then
        strStartDate = mstrBegin
        strEndDate = mstrEnd
    End If
End Function
