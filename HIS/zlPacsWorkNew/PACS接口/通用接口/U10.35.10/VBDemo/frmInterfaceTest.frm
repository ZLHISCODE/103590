VERSION 5.00
Begin VB.Form frmInterfaceTest 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   10230
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSendReportImageEx 
      Caption         =   "SendReportImageEx"
      Height          =   495
      Left            =   6690
      TabIndex        =   15
      Top             =   2955
      Width           =   1815
   End
   Begin VB.TextBox txtAdvice 
      Height          =   420
      Left            =   7035
      TabIndex        =   14
      Text            =   "1301"
      Top             =   360
      Width           =   1410
   End
   Begin VB.TextBox txtDepartId 
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Text            =   "46"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GetRequestInf1"
      Height          =   615
      Left            =   465
      TabIndex        =   10
      Top             =   4485
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SaveEcgReport"
      Height          =   735
      Left            =   3540
      TabIndex        =   9
      Top             =   3300
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetRequestInf"
      Height          =   615
      Left            =   435
      TabIndex        =   8
      Top             =   3660
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendReport 
      Caption         =   "SendReport"
      Height          =   615
      Left            =   6720
      TabIndex        =   7
      Top             =   1245
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsertReportAffix 
      Caption         =   "SendReportAffix"
      Height          =   375
      Left            =   6675
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton cmdInsertReportImage 
      Caption         =   "SendReportImage"
      Height          =   495
      Left            =   6675
      TabIndex        =   5
      Top             =   2205
      Width           =   1815
   End
   Begin VB.CommandButton cmdtest2 
      Caption         =   "GetRequestAdviceStatus"
      Height          =   375
      Left            =   465
      TabIndex        =   4
      Top             =   5250
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Connection"
      Height          =   375
      Left            =   2730
      TabIndex        =   3
      Top             =   390
      Width           =   2055
   End
   Begin VB.CommandButton cmdGetChargeTypes 
      Caption         =   "GetChargeTypes"
      Height          =   495
      Left            =   420
      TabIndex        =   1
      Top             =   1695
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetDeptItem 
      Caption         =   "GetDeptItem"
      Height          =   495
      Left            =   435
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetPatientByKey 
      Caption         =   "GetPatientByKey"
      Height          =   615
      Left            =   405
      TabIndex        =   2
      Top             =   2955
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "医嘱ID："
      Height          =   300
      Left            =   6210
      TabIndex        =   13
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "部门ID："
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmInterfaceTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iPacs As clsPacsInterface


Private Sub cmdGetChargeTypes_Click()
    
    Call iPacs.GetChargeTypes
    
    MsgBox UBound(iPacs.Tables.strDatas())
End Sub

Private Sub cmdGetDeptItem_Click()
    
    Call iPacs.GetDeptItems
    
    MsgBox UBound(iPacs.Tables.strDatas())
End Sub


Private Sub cmdGetPatientByKey_Click()
    Call iPacs.GetPatientInfo("22", pwtInHospital)
    
    MsgBox iPacs.GetCurValueByColumnName(0, "姓名")
End Sub

Private Sub cmdInsertReportAffix_Click()
    Call iPacs.SendReportAffix(Val(txtAdvice.Text), App.Path + "\1.bmp")
End Sub

Private Sub cmdInsertReportImage_Click()
    Call iPacs.SendReportImages(Val(txtAdvice.Text), App.Path + "\1.bmp")
    Call iPacs.GetLastError
End Sub




Private Sub cmdSendReport_Click()
    Call iPacs.DeleteReport(Val(txtAdvice.Text))
    Call iPacs.GetLastError
    
    Call iPacs.SendReport(Val(txtAdvice.Text), "测试我的报告发送。", "测试我的报告发送-诊断建议。", "X医生", "XXX")
    Call iPacs.GetLastError
End Sub

Private Sub cmdTest_Click()
    Set iPacs = New clsPacsInterface

'    iPacs.SplitChar = "*"
'    iPacs.NullValue = "NIL"
    If Not iPacs.InitInterface(Val(txtDepartId.Text), "local", "zlhis", "aqa", 100, "zlhis", "", "#", estShowMsg) Then
        Call iPacs.GetLastError
    End If
    
    Me.Caption = "Connected."
End Sub

Private Sub cmdtest2_Click()
    Call iPacs.GetRequestAdviceStatus(704)
    
    MsgBox iPacs.GetRecordCount(iPacs.Tables.strDatas)
End Sub

Private Sub Command1_Click()
    Call iPacs.GetRequestInfo("22", rwtInHospital)
    
    Dim strError As String
    strError = iPacs.GetLastError
    
    If strError <> "" Then
        MsgBox strError
        Exit Sub
    End If
    
    MsgBox iPacs.GetCurValueByColumnName(0, "姓名")
End Sub

Private Sub Command2_Click()
    Call iPacs.DeleteElectrocardioReport(8901)
    
    Call iPacs.SendElectrocardioReport(8901, "心电报告测试", _
        "c:\temp\scan.bmp", _
        "心电检查所见", "心电检查建议", "报告一声", "")
        
    Dim strError As String
    strError = iPacs.GetLastError
    
    If strError <> "" Then
        MsgBox strError
        Exit Sub
    End If
    
    MsgBox "ok"
        
End Sub

Private Sub Command3_Click()
    Call iPacs.GetRequestInfo1("2011-07-01", "2011-08-01", "心电")
    
    Dim strError As String
    strError = iPacs.GetLastError
    
    If strError <> "" Then
        MsgBox strError
        Exit Sub
    End If
    
    MsgBox iPacs.GetCurValueByColumnName(0, "姓名")
End Sub

