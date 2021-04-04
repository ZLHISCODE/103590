VERSION 5.00
Begin VB.Form frmLNCASet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6810
   Icon            =   "frmLNCASet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6810
      TabIndex        =   1
      Top             =   2865
      Width           =   6810
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Left            =   4425
         TabIndex        =   3
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   5625
         TabIndex        =   2
         Top             =   150
         Width           =   1100
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Index           =   2
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.CheckBox chkTS 
         BackColor       =   &H8000000E&
         Caption         =   "启用辽宁嘉鸿"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1845
         Width           =   1455
      End
      Begin VB.TextBox txtUrl 
         Height          =   360
         Left            =   1410
         TabIndex        =   5
         Top             =   765
         Width           =   5265
      End
      Begin VB.TextBox txtPenUrl 
         Height          =   360
         Left            =   1410
         TabIndex        =   4
         Top             =   1245
         Width           =   5265
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "服务器URL示例:http://218.25.86.214:2010/ssoworker"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   4410
      End
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "签名服务器URL"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   855
         Width           =   1170
      End
      Begin VB.Label lblPenUrl 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "手签服务器URL"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   1335
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmLNCASet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnOk As Boolean
    
    gudtPara.strSignURL = txtUrl.Text
    gudtPara.bytSignVersion = chkTS.Value
    gudtPara.strSIGNIP = txtPenUrl.Text
    
    gstrPara = LNCA_SetParaStr
    On Error GoTo errH
    strSQL = "Select count(1) as RowCount  From zlParameters Where 系统 = [1] And Nvl(模块, 0) = 0 And 参数号 = 90000"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "电子签名参数", glngSys)
    If Not rsTmp.EOF Then
        If rsTmp!RowCount = 0 Then
            lngID = gobjComLib.zlDatabase.GetNextId("zlParameters")
            strSQL = "Insert Into zlParameters(ID, 系统, 模块, 参数号, 参数名, 参数值) Values (" & lngID & ", " & glngSys & ", Null, 90000, '电子签名参数','" & gstrPara & "')"
            Call gobjComLib.zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
            blnOk = True
        End If
    End If
    
    If Not blnOk Then
        Call gobjComLib.zlDatabase.SetPara(90000, gstrPara, glngSys)
    End If
    
    Unload Me
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Load()
    Call LNCA_GetPara
    txtUrl.Text = gudtPara.strSignURL
    txtPenUrl.Text = gudtPara.strSIGNIP '辽宁CA没有用到这个存参数，暂时用这个存手签URL
    chkTS.Value = gudtPara.bytSignVersion
End Sub
