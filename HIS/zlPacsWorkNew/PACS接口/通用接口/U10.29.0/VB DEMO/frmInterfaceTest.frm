VERSION 5.00
Begin VB.Form frmInterfaceTest 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "GetRequestInf1"
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SaveEcgReport"
      Height          =   735
      Left            =   5640
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetRequestInf"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendReport 
      Caption         =   "SendReport"
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsertReportAffix 
      Caption         =   "InsertReportAffix"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdInsertReportImage 
      Caption         =   "InsertReportImage"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdtest2 
      Caption         =   "GetRequestAdviceStatus"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdGetChargeTypes 
      Caption         =   "GetChargeTypes"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetDeptItem 
      Caption         =   "GetDeptItem"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetPatientByKey 
      Caption         =   "GetPatientByKey"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
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
    Call iPacs.SendReportAffix(1901, "C:\Temp\midi.gif")
End Sub

Private Sub cmdInsertReportImage_Click()
    Call iPacs.SendReportImages(1901, "C:\Temp\Desert.jpg*c:\Temp\Scan.bmp")
    
    
End Sub

Private Sub cmdSendReport_Click()
    Call iPacs.DeleteReport(4255999)
    Call iPacs.SendReport(4255999, "测试我的报告发送。", "测试我的报告发送-诊断建议。", "涂医生", "XXX")
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
    Call iPacs.DeleteElectrocardioReport(3579381)
    
    Call iPacs.SendElectrocardioReport(3579381, "心电报告测试", _
        "E:\ZLHIS\ZLPacsWork\綦江心电接口\Ecg1.jpg#E:\ZLHIS\ZLPacsWork\綦江心电接口\Ecg2.jpg", _
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

Private Sub Form_Load()
    Set iPacs = New clsPacsInterface
    
'    iPacs.SplitChar = "*"
'    iPacs.NullValue = "NIL"
    If Not iPacs.InitInterface("", "ZLHIS", "aqa", 100, "ZLHIS", "", "#", estNoDisplay) Then
        Call iPacs.GetLastError
    End If
End Sub
