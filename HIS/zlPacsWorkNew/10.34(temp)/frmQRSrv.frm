VERSION 5.00
Begin VB.Form frmQrSrv 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraQueryRetrieve 
      Height          =   1935
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      Begin VB.Frame frmPatientID 
         Caption         =   "病人ID匹配"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   11055
         Begin VB.OptionButton optMatch 
            Caption         =   "医嘱ID"
            Height          =   195
            Index           =   2
            Left            =   8160
            TabIndex        =   5
            ToolTipText     =   "按医嘱ID将病人和接收的影像进行匹配"
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "病人标识号（门诊/住院号）"
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   4
            ToolTipText     =   "按病人标识号将病人和接收的影像进行匹配"
            Top             =   480
            Width           =   2610
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "检查号"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   3
            ToolTipText     =   "按检查号将病人和接收的影像进行匹配"
            Top             =   480
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkAcceptCGET 
         Caption         =   "支持C-GET"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmQrSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSrvID As Long
Public Sub ShowRefresh(ByVal SrvID As Long)
    mlngSrvID = SrvID
    If mlngSrvID = 0 Then
        fraQueryRetrieve.Caption = "上方列表中所选服务尚未保存，不能进行设置！"
        fraQueryRetrieve.Enabled = False
    Else
        fraQueryRetrieve.Caption = ""
        fraQueryRetrieve.Enabled = True
    End If
    RefreshPara
End Sub

Public Sub SavePara()
    Dim strData As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'支持C-GET','" & chkAcceptCGET.value & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存支持C-GET")
    
    strData = 0
    For i = 0 To optMatch.UBound
        If optMatch(i).value = True Then
            strData = i
            Exit For
        End If
    Next
    If strData = "" Then strData = "0"
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'病人ID匹配','" & strData & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病人ID匹配")
   Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshPara()
Dim rsTemp As New ADODB.Recordset, i As Integer
        gstrSQL = "select 服务ID,参数名称 ,参数值 from 影像DICOM服务参数 where 服务ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", mlngSrvID)
    chkAcceptCGET.value = False
    Do Until rsTemp.EOF
        Select Case rsTemp!参数名称
            Case "支持C-GET"
                chkAcceptCGET.value = Nvl(rsTemp!参数值)
            Case "病人ID匹配"
                optMatch(Nvl(rsTemp!参数值, 0)) = True
        End Select
        rsTemp.MoveNext
    Loop
End Sub

