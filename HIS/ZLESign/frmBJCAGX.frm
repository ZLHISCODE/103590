VERSION 5.00
Begin VB.Form frmBJCAGX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8070
   Icon            =   "frmBJCAGX.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8070
      TabIndex        =   10
      Top             =   2760
      Width           =   8070
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   5595
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   6795
         TabIndex        =   11
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Index           =   4
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "患者签名设置"
         Height          =   1530
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   7785
         Begin VB.TextBox txtdown 
            Height          =   270
            Left            =   2040
            TabIndex        =   9
            Text            =   "http://116.252.222.90:36069/verify/getVerifyData"
            Top             =   922
            Width           =   5640
         End
         Begin VB.TextBox txtUp 
            Height          =   270
            Left            =   2040
            TabIndex        =   8
            Text            =   "http://116.252.222.90:36069/verify/signature"
            Top             =   397
            Width           =   5640
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "患者签名获取URL"
            Height          =   225
            Left            =   315
            TabIndex        =   7
            Top             =   945
            Width           =   1425
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "患者签名上传URL"
            Height          =   225
            Left            =   315
            TabIndex        =   6
            Top             =   420
            Width           =   1425
         End
      End
      Begin VB.Frame fraPara 
         BackColor       =   &H80000005&
         Caption         =   "签名算法"
         Height          =   795
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7785
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "SM2"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   4
            Top             =   345
            Width           =   735
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "RSA"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   3
            Top             =   345
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CheckBox chkTS 
            BackColor       =   &H80000005&
            Caption         =   "启用签章"
            Height          =   255
            Index           =   5
            Left            =   4680
            TabIndex        =   2
            Top             =   345
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmBJCAGX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum CMD_ENUM
    CMD_OK = 0
    CMD_CANCEL = 1
End Enum

Private Sub chkTS_Click(Index As Integer)
    gudtPara.blnSignPic = chkTS(Index).Value = vbChecked
End Sub

Private Sub cmdPara_Click(Index As Integer)
    Dim objCA As New clsHNCA
    Dim lngID As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnOk As Boolean
    
    If Index = CMD_OK Then
        gudtPara.strSignURL = txtUp.Text & "|" & txtdown.Text
        gstrPara = BJCAGX_SetParaStr
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
    End If
    
    Unload Me
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Load()
    Call BJCAGX_GetPara
    opt(2).Value = gudtPara.bytSignVersion = V_RSA
    opt(3).Value = gudtPara.bytSignVersion = V_SM2
    chkTS(5).Value = IIf(gudtPara.blnSignPic, vbChecked, vbUnchecked)
    txtUp.Text = Split(gudtPara.strSignURL, "|")(0)
    If txtUp.Text = "" Then txtUp.Text = "http://116.252.222.90:36069/verify/signature"
    txtdown.Text = Split(gudtPara.strSignURL, "|")(1)
    If txtdown.Text = "" Then txtdown.Text = "http://116.252.222.90:36069/verify/getVerifyData"
End Sub

Private Sub opt_Click(Index As Integer)
    gudtPara.bytSignVersion = IIf(Index = 2, 0, 1)
End Sub
