VERSION 5.00
Begin VB.Form frmDrugPackerSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品分包机接口设置"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3210
   Icon            =   "frmDrugPackerSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3210
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "退出(&Q)"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "批量传送药品基础数据(&T)"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "分包机数据库连接设置(&S)"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmDrugPackerSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResult As String
Private mblnSetup As Boolean
Private mblnOutsideConnected As Boolean
Private mcnHIS As ADODB.Connection

Public Property Get ResultString() As String
    ResultString = mstrResult
End Property

Public Property Let ConnectHIS(ByVal cnHIS As ADODB.Connection)
    Set mcnHIS = cnHIS
End Property
Public Property Let OutsideConnected(ByVal blnConnected As Boolean)
    mblnOutsideConnected = blnConnected
End Property

Private Sub cmdConnect_Click()
    frmOutsideLinkSet.Show vbModal, Me
    mblnSetup = frmOutsideLinkSet.gblnSetupFinish
End Sub

Private Sub cmdQuit_Click()
    If mblnOutsideConnected Then
        mstrResult = "1"
    Else
        mstrResult = IIf(mblnSetup, "1", "0")
    End If
    Unload Me
End Sub

Private Sub cmdTrans_Click()
    Dim strTmp As String, strInsert As String, strDrugModel As String
    Dim rsTmp As New ADODB.Recordset, cmdInsert As New ADODB.Command
    Dim lngExec As Long

    If mcnHIS Is Nothing Or mcnHIS.State = adStateClosed Then
        MsgBox "ZLHIS数据库未连接！", vbCritical, GSTR_MESSAGE
        Exit Sub
    End If
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        MsgBox "你未连接药品分包数据库，请先执行DBConnect()函数！", , GSTR_MESSAGE
        Exit Sub
    End If
    
    '初始化lvwModel
    If frmSelModel.mcnHIS.ConnectionString = "" Then Set frmSelModel.mcnHIS = mcnHIS
    If frmSelModel.lvwModel.ColumnHeaders.Count <= 0 Then
        frmSelModel.InitLvwModel ' mcnHIS
    End If
    frmSelModel.Show vbModal
    strDrugModel = frmSelModel.DrugModel
End Sub
