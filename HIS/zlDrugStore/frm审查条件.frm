VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm审查条件 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "处方审查条件"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frm审查条件.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraUnit 
      Caption         =   "药品显示单位"
      ForeColor       =   &H00800000&
      Height          =   765
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   3165
      Begin VB.OptionButton optUnit 
         Caption         =   "药房单位"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "售价单位"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "处方类型"
      Height          =   930
      Left            =   90
      TabIndex        =   8
      Top             =   1320
      Width           =   3165
      Begin VB.CheckBox chk处方类型 
         Caption         =   "住院记帐处方"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   11
         Top             =   615
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox chk处方类型 
         Caption         =   "门诊记帐处方"
         Height          =   225
         Index           =   1
         Left            =   1725
         TabIndex        =   10
         Top             =   300
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox chk处方类型 
         Caption         =   "门诊收费处方"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   300
         Value           =   1  'Checked
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3570
      TabIndex        =   5
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   6
      Top             =   960
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "时间范围"
      Height          =   1155
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3165
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1260
         TabIndex        =   4
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   52625411
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1260
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   52625411
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "结束时间(&E)"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&B)"
         Height          =   180
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   3585
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
End
Attribute VB_Name = "frm审查条件"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mdatBegin As Date, mdatEnd As Date
Dim mintUnit As Integer
Dim mstrType As String          '处方类型(多选)
Dim mstrPrivs As String
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub





Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim n As Integer
    
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "开始时间大于结束时间了。", vbExclamation, gstrSysName
        dtpBegin.SetFocus
        Exit Sub
    End If
    
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    mintUnit = IIf(optUnit(0).Value = True, 0, 1)
    
    mstrType = ""
    For n = 0 To 2
        If chk处方类型(n).Value = 1 Then
            mstrType = mstrType & n
        End If
    Next
    
    If mstrType = "" Then
        MsgBox "请选择处方类型。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function GetCondition(datBegin As Date, datEnd As Date, intUnit As Integer, strType As String, strPrivs As String, ByVal frmOwner As Form) As Boolean
    dtpBegin.Value = datBegin
    dtpEnd.Value = datEnd
    dtpBegin.MaxDate = zldatabase.Currentdate
    dtpEnd.MaxDate = dtpBegin.MaxDate
    mstrPrivs = strPrivs
    
    If strType <> "" Then
        If InStr(1, strType, "0") > 0 Then
            chk处方类型(0).Value = 1
        Else
            chk处方类型(0).Value = 0
        End If
        If InStr(1, strType, "1") > 0 Then
            chk处方类型(1).Value = 1
        Else
            chk处方类型(1).Value = 0
        End If
        If InStr(1, strType, "2") > 0 Then
            chk处方类型(2).Value = 1
        Else
            chk处方类型(2).Value = 0
        End If
    End If
    
    If intUnit = 1 Then
        optUnit(1).Value = True
    Else
        optUnit(0).Value = True
    End If
    
    frm审查条件.Show vbModal, frmOwner
    GetCondition = mblnOK
    
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
        intUnit = mintUnit
        strType = mstrType
    End If
End Function

Private Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    objTxt.SelStart = 0
    objTxt.SelLength = Len(objTxt.Text)
End Sub
