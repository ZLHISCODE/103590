VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "日期选择"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPTime 
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   226557955
      CurrentDate     =   40832
   End
   Begin MSComCtl2.MonthView MonthView 
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   226557953
      CurrentDate     =   40832
   End
End
Attribute VB_Name = "frmDateSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrDate As String
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub cmdOK_Click()
    mstrDate = DTPTime
    Unload Me
End Sub

Private Sub DTPTime_Change()
    MonthView = DTPTime
End Sub

Private Sub DTPTime_DblClick()
    cmdOK_Click
End Sub

Private Sub Form_Load()
    MonthView = Now
    DTPTime = Now
End Sub

Private Sub MonthView_DblClick()
    cmdOK_Click
End Sub

Private Sub MonthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    DTPTime = Format(MonthView.value, "yyyy-MM-dd") & " " & Format(DTPTime, "HH:mm:ss")
End Sub
Public Function ShowMe(objFrm As Object) As String
    mstrDate = ""
    Me.Show vbModal, objFrm
    ShowMe = mstrDate
End Function
