VERSION 5.00
Begin VB.Form frmFeeGroupSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "轧帐缴款书打印方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   5895
      Begin VB.OptionButton optReportPrint 
         Caption         =   "选择打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Top             =   278
         Width           =   1215
      End
      Begin VB.OptionButton optReportPrint 
         Caption         =   "自动打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   278
         Width           =   1215
      End
      Begin VB.OptionButton optReportPrint 
         Caption         =   "不打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   278
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintReportSetting 
         Caption         =   "打印设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   7
         Top             =   315
         Width           =   1300
      End
   End
   Begin VB.Frame fraCollectPrint 
      Caption         =   "收款收据单打印方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   5895
      Begin VB.CommandButton cmdCollectPrintSetting 
         Caption         =   "打印设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   3
         Top             =   315
         Width           =   1300
      End
      Begin VB.OptionButton optCollectPrint 
         Caption         =   "选择打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   2
         Top             =   338
         Width           =   1215
      End
      Begin VB.OptionButton optCollectPrint 
         Caption         =   "自动打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   338
         Width           =   1215
      End
      Begin VB.OptionButton optCollectPrint 
         Caption         =   "不打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   338
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3720
      TabIndex        =   8
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4920
      TabIndex        =   10
      Top             =   3960
      Width           =   1100
   End
   Begin VB.Image imgSetting 
      Height          =   720
      Left            =   240
      Picture         =   "frmFeeGroupSetting.frx":0000
      Top             =   240
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   120
      X2              =   6120
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   120
      X2              =   6120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "请根据实际情况对参数进行设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1080
      TabIndex        =   13
      Top             =   660
      Width           =   2940
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "对收款收据单和轧帐缴款书的打印方式进行设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1080
      TabIndex        =   12
      Top             =   360
      Width           =   4410
   End
End
Attribute VB_Name = "frmFeeGroupSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnPrivs As Boolean, mlngModual As Long, mstrPrivs As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCollectPrintSetting_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_BILL_1507", Me)
End Sub

Private Sub cmdOK_Click()
    If optCollectPrint(0).Value = True Then
        zlDatabase.SetPara "收款收据打印方式", "0", glngSys, mlngModual, mblnPrivs
    End If
    If optCollectPrint(1).Value = True Then
        zlDatabase.SetPara "收款收据打印方式", "1", glngSys, mlngModual, mblnPrivs
    End If
    If optCollectPrint(2).Value = True Then
        zlDatabase.SetPara "收款收据打印方式", "2", glngSys, mlngModual, mblnPrivs
    End If
    If optReportPrint(0).Value = True Then
        zlDatabase.SetPara "缴款书打印方式", "0", glngSys, mlngModual, mblnPrivs
    End If
    If optReportPrint(1).Value = True Then
        zlDatabase.SetPara "缴款书打印方式", "1", glngSys, mlngModual, mblnPrivs
    End If
    If optReportPrint(2).Value = True Then
        zlDatabase.SetPara "缴款书打印方式", "2", glngSys, mlngModual, mblnPrivs
    End If
    Unload Me
End Sub

Private Sub cmdPrintReportSetting_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_INSIDE_1507", Me)
End Sub

Private Sub Form_Load()
    Dim i, j As Integer
    i = Val(zlDatabase.GetPara("收款收据打印方式", glngSys, mlngModual, "0", _
            Array(optCollectPrint(0), optCollectPrint(1), optCollectPrint(2)), mblnPrivs))
    If i < 0 Or i > 2 Then
        optCollectPrint(0).Value = True
    Else
        optCollectPrint(i).Value = True
    End If
    j = Val(zlDatabase.GetPara("缴款书打印方式", glngSys, mlngModual, "0", _
            Array(optReportPrint(0), optReportPrint(1), optReportPrint(2)), mblnPrivs))
    If j < 0 Or j > 2 Then
        optReportPrint(0).Value = True
    Else
        optReportPrint(j).Value = True
    End If
End Sub

Public Sub ParaSetting(frmMain As Object, lngModual As Long, strPrivs As String)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:外部调用接口
    '入参:frmMain-外部调用窗体
    '     lngModual-模块号
    '     mstrPrivs-权限串
    '编制:刘尔旋
    '日期:2013-09-22
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModual = lngModual
    mstrPrivs = strPrivs
    mblnPrivs = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    Me.Show vbModal, frmMain
End Sub

