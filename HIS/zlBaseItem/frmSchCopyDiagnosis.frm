VERSION 5.00
Begin VB.Form frmSchCopyDiagnosis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "复制预约设备的项目"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   ControlBox      =   0   'False
   Icon            =   "frmSchCopyDiagnosis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4095
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确定"
      Height          =   350
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1100
   End
   Begin VB.ComboBox cboItem 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "复制                 的全部诊疗项目"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3150
   End
End
Attribute VB_Name = "frmSchCopyDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDevice As Long
Private mstrResult As String


Public Function ShowMe(lngDevice As Long, ower As Object) As String
    mlngDevice = lngDevice
    mstrResult = ""
    
    Me.Show 1, ower
    
    ShowMe = mstrResult
    
End Function

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mstrResult = ""
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "提示"
    Err.Clear
End Sub

Private Sub cmdSure_Click()
    On Error GoTo errHandle
    
    If Len(cboItem.Text) = 0 Then
        MsgBox "请选择要复制的设备。", vbInformation, "提示"
        Exit Sub
    End If
    mstrResult = cboItem.Text
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "提示"
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    
    strSql = "Select 设备名称 From 影像预约设备 Where 影像类别 In (Select 影像类别 From 影像预约设备 Where Id = [1]) AND ID <> [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询预约设备", mlngDevice)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsTemp.EOF
        cboItem.AddItem rsTemp!设备名称
        rsTemp.MoveNext
    Loop
    
    If cboItem.ListCount > 0 Then cboItem.ListIndex = 0
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "提示"
    Err.Clear
End Sub
