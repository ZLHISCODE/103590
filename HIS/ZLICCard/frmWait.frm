VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "等待前置机响应"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4740
   Icon            =   "frmWait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2850
      Top             =   660
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   1
      Top             =   720
      Width           =   1100
   End
   Begin VB.Label lblnote 
      AutoSize        =   -1  'True
      Caption         =   "正在向前置机提交请求，请稍候......"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   3360
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintState As Integer    '0-未处理;1-正在处理;2-用户放弃;3-失败;9-成功
Private mstr日期 As String
Private mlng序列号 As Long
Private mstr错误信息 As String

'str错误信息：出错时记录错误信息；操作成功时记录接收到的数据
Public Function SendRequest(ByVal str日期 As String, ByVal lng序列号 As Long, ByRef str错误信息 As String) As Boolean
    mintState = 0
    mstr日期 = str日期
    mlng序列号 = lng序列号
    mstr错误信息 = ""
    Me.Show 1
    str错误信息 = mstr错误信息
    SendRequest = (mintState = 9)
End Function

Private Sub cmd取消_Click()
    '取消就直接退出就行了，不更新数据表中的状态
    mintState = 3
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    '检查数据是否已发送
    
    strSQL = " Select Nvl(标志,0) AS 标志,错误" & _
              " From 消息主表" & _
              " Where 日期='" & mstr日期 & "' And 序列号=" & mlng序列号
    Call OpenRecordset(rsTemp, "检查数据是否已发送", strSQL)
    mintState = rsTemp!标志
    mstr错误信息 = IIf(IsNull(rsTemp!错误), "", rsTemp!错误)
    
    Select Case mintState
    Case 0
        Exit Sub
    Case 1
        Me.lblnote.Caption = "正在向中心转发请求，请稍候......"
    Case 2
        Me.lblnote.Caption = "用户放弃！"
        Timer1.Enabled = False
    Case 3
        Me.lblnote.Caption = "发生错误，正在退出！"
        Timer1.Enabled = False
    Case 9
        Timer1.Enabled = False
        Call OrganizeData
    End Select
    
    If Timer1.Enabled = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    
    rsTemp.Open strSQL, gcnConnect, adOpenStatic, adLockReadOnly
    Set rsTemp.ActiveConnection = Nothing
End Sub

Private Sub OrganizeData()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = " Select 接收数据 From 消息接收 Where 日期='" & mstr日期 & "' And 序列号=" & mlng序列号 & " Order by 行号"
    Call OpenRecordset(rsTemp, "提取接收数据", strSQL)
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    With rsTemp
        mstr错误信息 = ""
        Do While Not .EOF
            mstr错误信息 = mstr错误信息 & IIf(IsNull(!接收数据), "", !接收数据)
            .MoveNext
        Loop
    End With
End Sub
