VERSION 5.00
Begin VB.Form frmInAdviceSendCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送选项"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmInAdviceSendCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraDetail 
      Height          =   1920
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   0
      Width           =   4500
      Begin VB.CheckBox chk加班加价 
         Caption         =   "执行加班加价(&V)"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   1620
         Width           =   1650
      End
      Begin VB.ListBox lstClass 
         Columns         =   4
         Height          =   1110
         ItemData        =   "frmInAdviceSendCond.frx":058A
         Left            =   195
         List            =   "frmInAdviceSendCond.frx":058C
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   450
         Width           =   4095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "要发送的类别(&T):"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   225
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3375
      TabIndex        =   4
      Top             =   2010
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2250
      TabIndex        =   3
      Top             =   2010
      Width           =   1100
   End
End
Attribute VB_Name = "frmInAdviceSendCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr类别s As String 'OUT:诊疗类别串
Public mblnOK As Boolean 'OUT:是否确认

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    '诊疗类别
    mstr类别s = ""
    For i = 0 To lstClass.ListCount - 1
        If lstClass.Selected(i) Then
            mstr类别s = mstr类别s & ",'" & Chr(lstClass.ItemData(i)) & "'"
        End If
    Next
    mstr类别s = Mid(mstr类别s, 2)
    If mstr类别s = "" Then
        MsgBox "请至少选择一种诊疗类别。", vbInformation, gstrSysName
        lstClass.SetFocus: Exit Sub
    End If
    If UBound(Split(mstr类别s, ",")) + 1 = lstClass.ListCount Then
        mstr类别s = ""
    End If
    
    gbln加班加价 = chk加班加价.Value = 1
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        j = lstClass.ListIndex
        For i = 0 To lstClass.ListCount - 1
            lstClass.Selected(i) = True
        Next
        lstClass.ListIndex = j
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        j = lstClass.ListIndex
        For i = 0 To lstClass.ListCount - 1
            lstClass.Selected(i) = False
        Next
        lstClass.ListIndex = j
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    chk加班加价.Value = IIF(gbln加班加价, 1, 0)
    '诊疗类别
    Call Load诊疗类别
End Sub

Private Function Load诊疗类别() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str类别s As String
    
    On Error GoTo errH
    
    str类别s = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "住院临嘱发送类别", "")
    
    strSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('4','7','9') Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstClass.AddItem rsTmp!名称
        lstClass.ItemData(lstClass.NewIndex) = Asc(rsTmp!编码)
        If str类别s <> "" Then
            If InStr(str类别s, "'" & rsTmp!编码 & "'") > 0 Then
                lstClass.Selected(lstClass.NewIndex) = True
            End If
        Else
            lstClass.Selected(lstClass.NewIndex) = True
        End If
        rsTmp.MoveNext
    Next
    If lstClass.ListCount > 0 Then lstClass.ListIndex = 0
    Load诊疗类别 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    '保存条件设置
    If mblnOK Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "住院临嘱发送类别", mstr类别s
    End If
End Sub
