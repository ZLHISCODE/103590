VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiseaseReportFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmDiseaseReportFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   2
      Top             =   3810
      Width           =   1200
   End
   Begin VB.Frame fraDate 
      Caption         =   "填报时间"
      Height          =   1785
      Left            =   195
      TabIndex        =   3
      Top             =   0
      Width           =   4530
      Begin VB.OptionButton optDates 
         Caption         =   "&1)最近的疾病报告(默认):"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2565
      End
      Begin VB.OptionButton optDates 
         Caption         =   "&2)指定日期范围的疾病报告:"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   690
         Width           =   2565
      End
      Begin VB.TextBox txtDates 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2805
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "7"
         Top             =   315
         Width           =   435
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   300
         Left            =   510
         TabIndex        =   4
         Top             =   990
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   67371011
         CurrentDate     =   38857
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   300
         Left            =   2175
         TabIndex        =   8
         Top             =   990
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   67371011
         CurrentDate     =   38857
      End
      Begin VB.Label lblDates 
         AutoSize        =   -1  'True
         Caption         =   "天"
         Height          =   180
         Left            =   3270
         TabIndex        =   9
         Top             =   375
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2325
      TabIndex        =   1
      Top             =   3810
      Width           =   1200
   End
   Begin VB.Frame fraFind 
      Caption         =   "查找条件"
      Height          =   1785
      Left            =   195
      TabIndex        =   0
      Top             =   1875
      Width           =   4530
      Begin VB.OptionButton optSearch 
         Caption         =   "报送人"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   307
         Width           =   900
      End
      Begin VB.TextBox txtInNo 
         Height          =   300
         Left            =   1710
         TabIndex        =   16
         Top             =   1365
         Width           =   1740
      End
      Begin VB.TextBox txtOutNo 
         Height          =   300
         Left            =   1710
         TabIndex        =   15
         Top             =   1005
         Width           =   1740
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1710
         TabIndex        =   14
         Top             =   645
         Width           =   1740
      End
      Begin VB.TextBox txtSender 
         Height          =   300
         Left            =   1710
         TabIndex        =   13
         Top             =   247
         Width           =   1740
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "住院号"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1425
         Width           =   900
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "门诊号"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1065
         Width           =   915
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "姓名"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   705
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmDiseaseReportFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mintDates As Integer, mstrDateFrom As String, mstrDateTo As String, mstrSender As String, mstrPatient As String, mlngOutNo As Long, mlngInNo As Long
Public Function ShowMe(ByVal frmobj As Object, ByRef intDates As Integer, ByRef strDateFrom As String, ByRef strDateTo As String, ByRef strSender As String, ByRef strPatient As String, ByRef lngOutNo As Long, ByRef lngInNo As Long) As Boolean
    If intDates <> 0 Then
        optDates(0).Value = True
        txtDates.Text = intDates
        dtpFrom.Value = Format(Now(), "yyyy-MM-dd")
        dtpFrom.Value = Format(Now() - intDates, "yyyy-MM-dd")
    Else
        optDates(1).Value = True
        dtpFrom.Value = Format(strDateFrom, "yyyy-MM-dd")
        dtpTo.Value = Format(strDateTo, "yyyy-MM-dd")
    End If
    optSearch(0).Value = True
    Me.Show 1, frmobj
    If mblnOk Then
        intDates = mintDates
        strDateFrom = mstrDateFrom
        strDateTo = mstrDateTo
        strSender = mstrSender
        strPatient = mstrPatient
        lngOutNo = mlngOutNo
        lngInNo = mlngInNo
        ShowMe = True
    Else
        ShowMe = False
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If optSearch(0).Value And Trim(txtPatient.Text) = "" Then
        MsgBox "未指定姓名,请检查!", vbInformation, gstrSysName
        Exit Sub
    ElseIf optSearch(1).Value And Trim(txtOutNo.Text) = "" Then
        MsgBox "未指定门诊号,请检查!", vbInformation, gstrSysName
        Exit Sub
    ElseIf optSearch(2).Value And Trim(txtInNo.Text) = "" Then
        MsgBox "未指定住院号,请检查!", vbInformation, gstrSysName
        Exit Sub
    ElseIf optSearch(3).Value And Trim(txtSender.Text) = "" Then
        MsgBox "未指定报送人,请检查!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mintDates = IIf(optDates(0).Value, Val(txtDates.Text), 0)
    mstrDateFrom = IIf(optDates(1).Value, Format(dtpFrom.Value, "yyyy-MM-dd"), "")
    mstrDateTo = IIf(optDates(1).Value, Format(dtpTo.Value, "yyyy-MM-dd"), "")
    mstrSender = IIf(optSearch(3).Value, txtSender.Text, "")
    mstrPatient = IIf(optSearch(0).Value, txtPatient.Text, "")
    mlngOutNo = IIf(optSearch(1).Value, Val(txtOutNo.Text), 0)
    mlngInNo = IIf(optSearch(2).Value, Val(txtInNo.Text), 0)
    mblnOk = True
    Unload Me
End Sub
Private Sub optDates_Click(Index As Integer)
    If Index = 0 Then
        txtDates.Enabled = True
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
    Else
        txtDates.Enabled = False
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
    End If
End Sub

Private Sub optSearch_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Select Case Index
        Case 3 '报送人
            txtSender.Enabled = True
            txtPatient.Enabled = False
            txtOutNo.Enabled = False
            txtInNo.Enabled = False
            optDates(0).Enabled = True
            optDates(1).Enabled = True
            Call txtSender.SetFocus
        Case 0 '姓名
            txtSender.Enabled = False
            txtPatient.Enabled = True
            txtOutNo.Enabled = False
            txtInNo.Enabled = False
            optDates(0).Enabled = False
            optDates(1).Enabled = False
            Call txtPatient.SetFocus
        Case 1 '门诊号
            txtSender.Enabled = False
            txtPatient.Enabled = False
            txtOutNo.Enabled = True
            txtInNo.Enabled = False
            optDates(0).Enabled = False
            optDates(1).Enabled = False
            Call txtOutNo.SetFocus
        Case 2 '住院号
            txtSender.Enabled = False
            txtPatient.Enabled = False
            txtOutNo.Enabled = False
            txtInNo.Enabled = True
            optDates(0).Enabled = False
            optDates(1).Enabled = False
            Call txtInNo.SetFocus
    End Select
End Sub

Private Sub txtInNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtInNo.Enabled = False Then
        optSearch(2).Value = True
    End If
End Sub

Private Sub txtOutNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtOutNo.Enabled = False Then
        optSearch(1).Value = True
    End If
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtPatient.Enabled = False Then
        optSearch(0).Value = True
    End If
End Sub

Private Sub txtSender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtSender.Enabled = False Then
        optSearch(3).Value = True
    End If
End Sub
