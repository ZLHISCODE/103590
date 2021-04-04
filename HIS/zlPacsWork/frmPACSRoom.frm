VERSION 5.00
Begin VB.Form frmPACSRoom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "执行间设置"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmPacsRoom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2250
      TabIndex        =   3
      Top             =   1560
      Width           =   1100
   End
   Begin VB.CheckBox chkOnly 
      Caption         =   "只处理安排到当前执行间的检查项目(&J)"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox cboRoom 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   345
      Width           =   2925
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3405
      TabIndex        =   4
      Top             =   1560
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "执行间(&S)"
      Height          =   350
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   6120
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前执行间(&R)"
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   405
      Width           =   1170
   End
End
Attribute VB_Name = "frmPACSRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrRoom As String, blnIfOnlyShow As Boolean, mlngDeptID As Long

Public Function ShowMe(objParent As Object, ByRef strRoom As String, ByRef ifOnlyShow As Boolean, _
    Optional lngDeptID As Long = 0) As Boolean
    mstrRoom = strRoom: blnIfOnlyShow = ifOnlyShow: mlngDeptID = lngDeptID
    
    On Local Error Resume Next
    Me.Show 1, objParent
    On Error GoTo 0
    
    strRoom = mstrRoom: ifOnlyShow = blnIfOnlyShow
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errH
    mstrRoom = cboRoom.Text: blnIfOnlyShow = IIf(chkOnly.Value = 0, False, True)
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前执行间", mstrRoom
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只处理当前执行间项目", blnIfOnlyShow
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSetup_Click()
    With frmTechnicRoom
        .lblDept.Tag = mlngDeptID
        .lblDept.Caption = "执行间"
        .Show 1, Me
        
        InitRoom
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    
    On Error GoTo errH
    InitRoom
    If cboRoom.ListCount = 0 Then
        If MsgBox("当前科室还没有设置执行间，是否设置?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Unload Me: Exit Sub
        Else
            cmdSetup_Click
        End If
    End If
    
    chkOnly.Value = IIf(blnIfOnlyShow, 1, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitRoom()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    On Error GoTo errH
    '执行间内容
    If mlngDeptID = 0 Then
        strSql = "Select * From 医技执行房间"
    Else
        strSql = "Select * From 医技执行房间 Where 科室ID= [1] "
    End If
    
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    cboRoom.Clear
    For i = 1 To rsTmp.RecordCount
        cboRoom.AddItem rsTmp!执行间
        rsTmp.MoveNext
    Next
    
    On Error Resume Next
    cboRoom.ListIndex = 0
    cboRoom.Text = mstrRoom
    cmdOK.Enabled = (cboRoom.ListCount > 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
