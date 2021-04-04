VERSION 5.00
Begin VB.Form FrmReqInput 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   Icon            =   "FrmReqInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "登记时光标不经过以下项目"
      Height          =   6450
      Left            =   80
      TabIndex        =   20
      Top             =   120
      Width           =   7725
      Begin VB.CheckBox ChkInput 
         Caption         =   "检查技师二"
         Height          =   180
         Index           =   24
         Left            =   6360
         TabIndex        =   25
         Top             =   3240
         Width           =   1260
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "附加主述"
         Height          =   180
         Index           =   23
         Left            =   420
         TabIndex        =   24
         Top             =   4080
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "检查技师"
         Height          =   180
         Index           =   22
         Left            =   5040
         TabIndex        =   23
         Top             =   3240
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "造影剂"
         Height          =   180
         Index           =   21
         Left            =   3765
         TabIndex        =   22
         Top             =   3240
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "出生日期"
         Height          =   180
         Index           =   20
         Left            =   3765
         TabIndex        =   21
         Top             =   486
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "检查时间"
         Height          =   180
         Index           =   19
         Left            =   2655
         TabIndex        =   19
         Top             =   3240
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "申请时间"
         Height          =   180
         Index           =   18
         Left            =   1545
         TabIndex        =   18
         Top             =   3240
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "紧急"
         Height          =   180
         Index           =   15
         Left            =   5040
         TabIndex        =   15
         Top             =   2322
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   $"FrmReqInput.frx":000C
         Height          =   180
         Index           =   17
         Left            =   420
         TabIndex        =   17
         Top             =   3240
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "检查设备"
         Height          =   180
         Index           =   16
         Left            =   6360
         TabIndex        =   16
         Top             =   2322
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "执行间"
         Height          =   180
         Index           =   14
         Left            =   3765
         TabIndex        =   14
         Top             =   2322
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "地址"
         Height          =   180
         Index           =   13
         Left            =   2655
         TabIndex        =   13
         Top             =   2322
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "邮编"
         Height          =   180
         Index           =   12
         Left            =   1545
         TabIndex        =   12
         Top             =   2322
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "电话"
         Height          =   180
         Index           =   11
         Left            =   420
         TabIndex        =   11
         Top             =   2322
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "婚姻"
         Height          =   180
         Index           =   10
         Left            =   3765
         TabIndex        =   10
         Top             =   1404
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "职业"
         Height          =   180
         Index           =   9
         Left            =   2655
         TabIndex        =   9
         Top             =   1404
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "民族"
         Height          =   180
         Index           =   8
         Left            =   6360
         TabIndex        =   8
         Top             =   1404
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "身份证号"
         Height          =   180
         Index           =   7
         Left            =   5040
         TabIndex        =   7
         Top             =   1404
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "付款方式"
         Height          =   180
         Index           =   6
         Left            =   5040
         TabIndex        =   6
         Top             =   486
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "费别"
         Height          =   180
         Index           =   5
         Left            =   6360
         TabIndex        =   5
         Top             =   486
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "体重"
         Height          =   180
         Index           =   4
         Left            =   1545
         TabIndex        =   4
         Top             =   1404
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "身高"
         Height          =   180
         Index           =   3
         Left            =   420
         TabIndex        =   3
         Top             =   1404
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "年龄"
         Height          =   180
         Index           =   2
         Left            =   2655
         TabIndex        =   2
         Top             =   486
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "性别"
         Height          =   180
         Index           =   1
         Left            =   1545
         TabIndex        =   1
         Top             =   486
         Width           =   1020
      End
      Begin VB.CheckBox ChkInput 
         Caption         =   "英文名"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   0
         Top             =   486
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FrmReqInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngDeptID As Long     '记录当前科室ID
Public mintType As Integer    '控制类型
Private mblnRefreshed As Boolean    '记录是否刷新

Private Sub Form_Activate()
    If mintType = 0 Then
        Frame1.Caption = "登记时光标不经过以下项目"
    Else
        Frame1.Caption = "登记时必须录入以下项目"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Debug.Print KeyAscii
End Sub

Public Sub zlRefresh()
    Dim i As Integer, strInput As String, j As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    mblnRefreshed = True
    
    '初始化选择框
    For i = 0 To ChkInput.UBound
        ChkInput(i).value = 0
    Next
    
    strSql = "select ID ,科室ID,参数值 from 影像流程参数 where 科室ID = [1] and 参数名 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID, CStr(IIf(mintType = 0, "输入控制", "必录控制")))
    
    If Not rsTemp.EOF Then
        strInput = Nvl(rsTemp!参数值)
        For i = 0 To UBound(Split(strInput, "|"))
        For j = 0 To ChkInput.UBound
            If ChkInput(j).Caption = Split(strInput, "|")(i) Then ChkInput(j).value = 1: Exit For
        Next
    Next
    End If
End Sub

Public Sub zlSave()
    Dim i As Integer, strInput As String
    Dim strSql As String
    
    If mblnRefreshed = False Then Exit Sub      '没有刷新则不保存
    
    For i = 0 To ChkInput.UBound
        If ChkInput(i).value = 1 Then strInput = strInput & "|" & ChkInput(i).Caption
    Next
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '" & IIf(mintType = 0, "输入控制", "必录控制") & "','" & strInput & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
End Sub

Private Sub Form_Load()
    mblnRefreshed = False
End Sub

Private Sub Form_Resize()
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
End Sub
