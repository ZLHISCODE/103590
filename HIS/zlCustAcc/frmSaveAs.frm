VERSION 5.00
Begin VB.Form frmSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "另存记帐单模板"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmSaveAs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSave 
      Caption         =   "新记帐单模板信息"
      Height          =   1485
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4395
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   960
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "名称"
         Top             =   840
         Width           =   3165
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   960
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "编码"
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3390
      TabIndex        =   6
      Top             =   1890
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2190
      TabIndex        =   5
      Top             =   1890
      Width           =   1100
   End
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrID As String
Dim mstr编码 As String
Dim mstr名称 As String
Dim mblnOK As Boolean
Dim mblnSave  As Boolean       '输入编码名称后，是否需要立即保存
Dim mblnChange As Boolean     '是否改变了

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save记帐单() = False Then Exit Sub
    
    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关记帐单的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 1
        strTemp = Trim(txtEdit(i).Text)
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "所输入内容不能超过" & Int(txtEdit(i).MaxLength / 2) & "个汉字" & "或" & txtEdit(i).MaxLength & "个字母。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    If Len(Trim(txtEdit(0).Text)) = 0 Then
        txtEdit(0).Text = ""
        MsgBox "编码不能为空。", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(1).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(1).Text = ""
        txtEdit(1).SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save记帐单() As Boolean
'功能:保存编辑的内容到记帐单表中
'参数:
'返回值:成功返回True,否则为False
    Dim lngID As Long
    On Error GoTo errHandle
    
    lngID = zlDatabase.GetNextId("收费记帐单")
    mstr编码 = txtEdit(0).Text
    mstr名称 = txtEdit(1).Text
    
    If mblnSave = True Then
        gstrSQL = "zl_收费记帐单_SaveAs('" & lngID & _
            "','" & mstr编码 & "','" & mstr名称 & "','" & mstrID & "')"
            
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
        
    mblnOK = True
    mblnChange = False
    mstrID = CStr(lngID)
    Save记帐单 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 另存模板(strID As String, str编码 As String, str名称 As String, Optional ByVal blnSave As Boolean = True) As Boolean
'功能:用来与调用的记帐单管理窗口进行通讯的程序
'参数:实际上是做为返回值
'     strID 在传入时是参照模板的ID，返回时是作为新增模板的ID
    mblnChange = False
    mblnSave = blnSave
    mblnOK = False
    mstrID = strID
    frmSaveAs.Show vbModal
    
    If mblnOK = True Then
        strID = mstrID
        str编码 = mstr编码
        str名称 = mstr名称
    End If
    另存模板 = mblnOK
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
    End If
End Sub


