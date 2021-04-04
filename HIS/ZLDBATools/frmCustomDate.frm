VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自定义时间选择"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4785
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSelected 
      Height          =   1785
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin MSComCtl2.MonthView monthView 
      Height          =   2370
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   221970433
      CurrentDate     =   43573
   End
   Begin VB.Label lblThree 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "添加近三周同日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3240
      MouseIcon       =   "frmCustomDate.frx":6852
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   1440
   End
   Begin VB.Line line 
      X1              =   0
      X2              =   4920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblDelete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "移除最近选中日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3240
      MouseIcon       =   "frmCustomDate.frx":69A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label lbClear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "清空已选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3960
      MouseIcon       =   "frmCustomDate.frx":6AF6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lblSelectd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已选日期"
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期选择器"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmCustomDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrResult As String

Public Function ShowSelector() As String
    Me.Show 1
    ShowSelector = mstrResult
    mstrResult = ""
End Function

Private Sub cmdOK_Click()
    mstrResult = txtSelected.Text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    monthView.Value = Date
End Sub

Private Sub lbClear_Click()
    txtSelected.Text = ""
End Sub

Private Sub lblDelete_Click()
    On Error Resume Next
    If txtSelected.Text <> "" And InStr(1, txtSelected.Text, vbNewLine) = 0 Then txtSelected.Text = ""
    txtSelected.Text = Mid(txtSelected.Text, 1, InStrRev(txtSelected.Text, vbNewLine) - 1)
End Sub

Private Sub lblThree_Click()
    Dim dtTmp As Date, arrTmp() As String
    Dim i As Integer
    
    On Error Resume Next
    If txtSelected.Text <> "" Then
        arrTmp = Split(txtSelected.Text, vbNewLine)
        dtTmp = CDate(arrTmp(UBound(arrTmp)))
    Else
        dtTmp = monthView.Value
    End If
    
    For i = 0 To 2
        txtSelected.Text = txtSelected.Text & IIf(txtSelected.Text = "", "", vbNewLine) & (dtTmp - 7 * i)
    Next
End Sub

Private Sub monthView_DateDblClick(ByVal DateDblClicked As Date)
    If txtSelected.Text = "" Then
        txtSelected.Text = DateDblClicked
    Else
        txtSelected.Text = txtSelected.Text & vbNewLine & DateDblClicked
    End If
End Sub



