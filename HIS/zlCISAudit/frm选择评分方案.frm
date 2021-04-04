VERSION 5.00
Begin VB.Form frm选择评分方案 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   Picture         =   "frm选择评分方案.frx":0000
   ScaleHeight     =   1920
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmbSelFA 
      Height          =   300
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   3075
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "导入(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   225
      TabIndex        =   1
      Top             =   1305
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   2085
      TabIndex        =   0
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "请选择导入的方案名称："
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   450
      Width           =   2400
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "导入方案"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "frm选择评分方案"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ID_From          As Long     '选中的源方案ID
Private ID()            As Long     '供选择的ID序列，与CmbBox对应

'==============================================================================
'=功能：取消退出
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：确定选中方案
'==============================================================================
Private Sub CmdOK_Click()
    On Error GoTo ErrH
    If cmbSelFA.ListIndex = -1 Then MsgBox "请选择一个方案供导入！", vbOKOnly + vbInformation, gstrSysName: Exit Sub
    ID_From = ID(cmbSelFA.ListIndex + 1)
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：填充方案选择下拉框
'=参数：传入将导入的ID号
'==============================================================================
Public Sub FillCmbSelFA(ID_to As Long)
    Dim rsTemp      As ADODB.Recordset
    Dim i           As Long
    
    On Error GoTo ErrH
    
    cmbSelFA.Clear
    '注意调用格式：先赋值gstrSQL,然后打开数据集
    gstrSQL = "select ID,名称,选用,启用时间 from 病案评分方案 where 类型='住院' and ID <> [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, ID_to)
    rsTemp.Sort = "选用 desc,名称 ,启用时间"
    
    i = 0
    Do Until rsTemp.EOF
        i = i + 1
        ReDim Preserve ID(1 To i) As Long
        cmbSelFA.AddItem rsTemp("名称"), i - 1
        ID(i) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    If i >= 1 Then cmbSelFA.ListIndex = 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

