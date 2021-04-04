VERSION 5.00
Begin VB.Form frmVsfColsList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据列表配置"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox lstVsfColsName 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      ItemData        =   "frmVsfColsList.frx":0000
      Left            =   120
      List            =   "frmVsfColsList.frx":0007
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "恢复默认(&D)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   3960
      Width           =   820
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&S)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "勾选需要显示的列"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmVsfColsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mStrDefaultCfg As String '默认配置
Private mStrCfgOld As String '当前使用的配置
Private mStrCfgNew As String '修改后配置

Public Function GetListCfg() As String
'获取更新后的配置串,若新旧相同，返回 ""
    GetListCfg = IIf(mStrCfgOld = mStrCfgNew, "", mStrCfgNew)
End Function


Public Sub ShowVsfColsListWindow(ByVal StrDefaultCfg As String, ByVal StrNowCfg As String, ByVal frmParent As Object)
'打开列名列表窗体并加载默认显示的列
    
    mStrDefaultCfg = StrDefaultCfg
    mStrCfgOld = StrNowCfg
    mStrCfgNew = StrNowCfg
    
    '调用加载默认显示的列
    Call LoadColsList(StrNowCfg)
    
    '加载窗体
    Me.Move frmParent.Left + (frmParent.Width - Me.Width) / 2, frmParent.Top + (frmParent.Height - Me.Height) / 2
    Call Show(1, frmParent)
End Sub



Private Sub CmdOK_Click()
'确定保存属性操作
On Error GoTo errHandle

    Dim i As Integer
    Dim strName As String
    Dim lngWidth As Long
    Dim intShow As Integer '1 显示    0 不显示
    Dim j As Integer
    Dim strCol() As String '每一个列信息
    
'    lngWidth = 800
    
    strCol = Split(mStrCfgNew, "|")
    mStrCfgNew = ""
    For i = 0 To UBound(strCol)
        If Len(mStrCfgNew) > 0 Then mStrCfgNew = mStrCfgNew & "|"
        
        strName = ""
        intShow = 1
        For j = 0 To lstVsfColsName.ListCount - 1
            strName = lstVsfColsName.List(j)
            intShow = IIf(lstVsfColsName.Selected(j), 1, 0)
            lngWidth = Split(strCol(i), ",")(1)
            
            If strName = Split(strCol(i), ",")(0) Then
                strCol(i) = strName & "," & lngWidth & "," & intShow
                Exit For
            End If
        Next
        
        mStrCfgNew = mStrCfgNew & strCol(i)
    Next
    
    Call Me.Hide
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadColsList(ByVal StrCfgOld As String)
'加载默认显示的列
On Error GoTo errHandle
    Dim i As Integer
    Dim lngUbound As Long
    Dim strValue As String

    
    lngUbound = UBound(Split(StrCfgOld, "|"))
    For i = 0 To lngUbound
        strValue = Split(StrCfgOld, "|")(i)
        lstVsfColsName.List(i) = Split(strValue, ",")(0)
        lstVsfColsName.Selected(i) = IIf(Split(strValue, ",")(2), True, False)
    Next

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdDefault_Click()
'恢复默认勾选:按照方案配置顺序显示，并且宽度一律为800，全部显示
On Error GoTo errHandle

    mStrCfgNew = mStrDefaultCfg
    Call LoadColsList(mStrDefaultCfg)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
'卸载窗体
On Error GoTo errHandle

    mStrCfgNew = mStrCfgOld
    Unload Me
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    '将窗口置顶
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3

End Sub

