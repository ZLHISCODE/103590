VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5145
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtTmpImgPath 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtReportID 
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtAdviceID 
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton GetReportList 
      Caption         =   "GetReportList"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton GetReportFormHandle 
      Caption         =   "GetReportFormHandle"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton ShowImage 
      Caption         =   "ShowImage"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtPatID 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton GetReportImage 
      Caption         =   "GetReportImage"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtPageID 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "缓存目录"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "医嘱ID"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "主页ID"
      Height          =   180
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "病人ID"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   540
   End
   Begin VB.Label 医嘱ID 
      AutoSize        =   -1  'True
      Caption         =   "报告ID"
      Height          =   180
      Left            =   2760
      TabIndex        =   6
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mobjPublicPACS As Object 'zlPublicPACS.clsPublicPacs

Private Sub Form_Load()
    Dim cnOracle As ADODB.Connection
    
    Dim strServer As String
    Dim strUser As String
    Dim strPwd As String
    
    strServer = "zlhis"
    strUser = "zlhis"
    strPwd = "aqa"
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    Set cnOracle = OraDataOpen(strServer, strUser, strPwd)
    
    If cnOracle Is Nothing Then Exit Sub
    
    Set mobjPublicPACS = CreateObject("zlPublicPACS.clsPublicPacs")
    Call mobjPublicPACS.InitInterface(cnOracle, UCase(strUser))
End Sub

Private Sub GetReportImage_Click()
    Dim objImgFileName As Collection
    
    If Trim(txtAdviceID.Text) = "" Then
        MsgBox "请输入医嘱ID"
        txtAdviceID.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTmpImgPath.Text) = "" Then
        MsgBox "请输入图像缓存目录"
        txtTmpImgPath.SetFocus
        Exit Sub
    End If
    
    Set objImgFileName = mobjPublicPACS.GetReportImage(txtAdviceID.Text, txtTmpImgPath.Text)
    
    If objImgFileName Is Nothing Then Exit Sub
    
    MsgBox "报告图象数量:" & objImgFileName.Count
End Sub

Private Sub GetReportList_Click()
    Dim rsData As ADODB.Recordset
    
    If Trim(txtPatID.Text) = "" Then
        MsgBox "请输入病人ID"
        txtPatID.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPageID.Text) = "" Then
        MsgBox "请输入主页ID"
        txtPageID.SetFocus
        Exit Sub
    End If
    
    Set rsData = mobjPublicPACS.GetReportList(txtPatID.Text, txtPageID.Text)
    
    MsgBox rsData.RecordCount
End Sub

Private Sub GetReportFormHandle_Click()
    Dim Pane1 As Pane
    Dim lngHandle As Long
    
    If Trim(txtReportID.Text) = "" Then
        MsgBox "请输入报告ID"
        txtReportID.SetFocus
        Exit Sub
    End If
    
    lngHandle = mobjPublicPACS.GetReportFormHandle(txtReportID.Text)
    
    frmMain.ShowMe lngHandle
End Sub

Private Sub ShowImage_Click()
    If Trim(txtAdviceID.Text) = "" Then
        MsgBox "请输入医嘱ID"
        txtAdviceID.SetFocus
        Exit Sub
    End If
    
    Call mobjPublicPACS.ShowImage(txtAdviceID.Text, Me, False)
End Sub
